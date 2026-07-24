using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using MiniWord.Models;
using MiniWord.Services;

namespace MiniWord
{
    public partial class MainWindow
    {
        private bool _showMarks;
        private bool _painterActive;
        private readonly List<(DependencyProperty prop, object value)> _painterFormat = new();
        private FindReplaceDialog? _findDialog;
        private DispatcherTimer? _autosaveTimer;
        private bool _draftDirty;

        private static readonly string AutosaveDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "MiniWord", "autosave");

        private void InitializeFeatures()
        {
            DataObject.AddPastingHandler(TextEditor, OnPaste);

            _autosaveTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(30) };
            _autosaveTimer.Tick += (s, e) => AutosaveTick();
            _autosaveTimer.Start();

            UpdateWordCount();
        }

        #region Clipboard

        // Only intercept pure image payloads; leave text/rtf to the default paste
        private void OnPaste(object sender, DataObjectPastingEventArgs e)
        {
            var data = e.DataObject;
            if (data.GetDataPresent(DataFormats.UnicodeText) || data.GetDataPresent(DataFormats.Text) || data.GetDataPresent(DataFormats.Rtf))
                return;

            if (data.GetDataPresent(DataFormats.Bitmap) && data.GetData(DataFormats.Bitmap) is BitmapSource src)
            {
                e.CancelCommand();
                InsertImageSource(src, src.Width, src.Height);
            }
        }

        private void MenuPasteText_Click(object sender, RoutedEventArgs e)
        {
            if (Clipboard.ContainsText())
            {
                TextEditor.Selection.Text = Clipboard.GetText();
                TextEditor.CaretPosition = TextEditor.Selection.End;
                _hasUnsavedChanges = true;
            }
            TextEditor.Focus();
        }

        private void InsertImageSource(ImageSource src, double naturalW, double naturalH)
        {
            double maxWidth = _pageSize.WidthDip - TextEditor.Padding.Left - TextEditor.Padding.Right;
            double w = naturalW, h = naturalH;
            if (w > maxWidth && w > 0)
            {
                h *= maxWidth / w;
                w = maxWidth;
            }
            var image = new Image { Source = src, Width = w, Height = h, Stretch = Stretch.Uniform };
            _ = new InlineUIContainer(image, TextEditor.CaretPosition);
            _hasUnsavedChanges = true;
            TextEditor.Focus();
        }

        #endregion

        #region Character formatting extras

        private void ToggleTextDecoration(TextDecoration deco)
        {
            var current = TextEditor.Selection.GetPropertyValue(Inline.TextDecorationsProperty) as TextDecorationCollection;
            var next = new TextDecorationCollection();
            bool present = false;
            if (current != null)
            {
                foreach (var d in current)
                {
                    if (d.Location == deco.Location)
                        present = true;
                    else
                        next.Add(d);
                }
            }
            if (!present)
                next.Add(deco);
            TextEditor.Selection.ApplyPropertyValue(Inline.TextDecorationsProperty, next);
            _hasUnsavedChanges = true;
            UpdateToolbarState();
            TextEditor.Focus();
        }

        private void StrikeButton_Click(object sender, RoutedEventArgs e) =>
            ToggleTextDecoration(TextDecorations.Strikethrough[0]);

        private void Superscript_Click(object sender, RoutedEventArgs e)
        {
            EditingCommands.ToggleSuperscript.Execute(null, TextEditor);
            _hasUnsavedChanges = true;
            UpdateToolbarState();
            TextEditor.Focus();
        }

        private void Subscript_Click(object sender, RoutedEventArgs e)
        {
            EditingCommands.ToggleSubscript.Execute(null, TextEditor);
            _hasUnsavedChanges = true;
            UpdateToolbarState();
            TextEditor.Focus();
        }

        #endregion

        #region Change case

        private void ChangeCase_Click(object sender, RoutedEventArgs e)
        {
            CaseMenu.PlacementTarget = ChangeCaseButton;
            CaseMenu.IsOpen = true;
        }

        private void TransformSelectedText(Func<string, string> transform)
        {
            var sel = TextEditor.Selection;
            if (sel.IsEmpty)
                return;
            var text = sel.Text;
            if (string.IsNullOrEmpty(text))
                return;
            sel.Text = transform(text);
            _hasUnsavedChanges = true;
            TextEditor.Focus();
        }

        private void CaseUpper_Click(object sender, RoutedEventArgs e) =>
            TransformSelectedText(t => t.ToUpper(CultureInfo.CurrentCulture));

        private void CaseLower_Click(object sender, RoutedEventArgs e) =>
            TransformSelectedText(t => t.ToLower(CultureInfo.CurrentCulture));

        private void CaseSentence_Click(object sender, RoutedEventArgs e) =>
            TransformSelectedText(ToSentenceCase);

        private void CaseTitle_Click(object sender, RoutedEventArgs e) =>
            TransformSelectedText(ToTitleCase);

        private static string ToSentenceCase(string text)
        {
            var chars = text.ToLower(CultureInfo.CurrentCulture).ToCharArray();
            bool capitalizeNext = true;
            for (int i = 0; i < chars.Length; i++)
            {
                if (char.IsLetter(chars[i]))
                {
                    if (capitalizeNext)
                    {
                        chars[i] = char.ToUpper(chars[i], CultureInfo.CurrentCulture);
                        capitalizeNext = false;
                    }
                }
                else if (chars[i] == '.' || chars[i] == '!' || chars[i] == '?')
                {
                    capitalizeNext = true;
                }
            }
            return new string(chars);
        }

        private static string ToTitleCase(string text)
        {
            var chars = text.ToCharArray();
            bool startOfWord = true;
            for (int i = 0; i < chars.Length; i++)
            {
                if (char.IsWhiteSpace(chars[i]))
                {
                    startOfWord = true;
                }
                else
                {
                    chars[i] = startOfWord
                        ? char.ToUpper(chars[i], CultureInfo.CurrentCulture)
                        : char.ToLower(chars[i], CultureInfo.CurrentCulture);
                    startOfWord = false;
                }
            }
            return new string(chars);
        }

        #endregion

        #region Format painter

        private static readonly DependencyProperty[] PainterProps =
        {
            TextElement.FontFamilyProperty, TextElement.FontSizeProperty,
            TextElement.FontWeightProperty, TextElement.FontStyleProperty,
            TextElement.ForegroundProperty, TextElement.BackgroundProperty,
            Inline.TextDecorationsProperty,
        };

        private void FormatPainter_Click(object sender, RoutedEventArgs e)
        {
            if (FormatPainterButton.IsChecked == true)
            {
                _painterFormat.Clear();
                foreach (var p in PainterProps)
                {
                    var v = TextEditor.Selection.GetPropertyValue(p);
                    if (v != DependencyProperty.UnsetValue && v != null)
                        _painterFormat.Add((p, v));
                }
                _painterActive = true;
                TextEditor.PreviewMouseLeftButtonUp += Painter_MouseUp;
            }
            else
            {
                StopFormatPainter();
            }
            TextEditor.Focus();
        }

        private void Painter_MouseUp(object sender, MouseButtonEventArgs e)
        {
            // Apply after the selection has settled
            Dispatcher.BeginInvoke(new Action(() =>
            {
                if (!_painterActive || TextEditor.Selection.IsEmpty)
                    return;
                foreach (var (prop, value) in _painterFormat)
                    TextEditor.Selection.ApplyPropertyValue(prop, value);
                _hasUnsavedChanges = true;
                StopFormatPainter();
            }), DispatcherPriority.Input);
        }

        private void StopFormatPainter()
        {
            _painterActive = false;
            FormatPainterButton.IsChecked = false;
            TextEditor.PreviewMouseLeftButtonUp -= Painter_MouseUp;
        }

        #endregion

        #region Line spacing & paragraph marks

        private IEnumerable<Paragraph> GetSelectedParagraphs()
        {
            var start = TextEditor.Selection.Start.Paragraph;
            var end = TextEditor.Selection.End.Paragraph;
            if (start == null)
                yield break;
            for (Block? b = start; b != null; b = b.NextBlock)
            {
                if (b is Paragraph p)
                    yield return p;
                if (b == end)
                    break;
            }
        }

        private void LineSpacing_Click(object sender, RoutedEventArgs e)
        {
            LineSpacingMenu.PlacementTarget = LineSpacingButton;
            LineSpacingMenu.IsOpen = true;
        }

        private void LineSpacingPreset_Click(object sender, RoutedEventArgs e)
        {
            if (sender is MenuItem item && item.Tag is string tag
                && double.TryParse(tag, NumberStyles.Any, CultureInfo.InvariantCulture, out var mult))
            {
                foreach (var p in GetSelectedParagraphs())
                {
                    double fontSize = p.FontSize > 0 ? p.FontSize : TextEditor.FontSize;
                    // 1.0 ≈ single line; use the common ~1.2 line box as the base
                    p.LineHeight = mult * fontSize * 1.2;
                }
                _hasUnsavedChanges = true;
                TextEditor.Focus();
            }
        }

        private void ShowMarks_Click(object sender, RoutedEventArgs e)
        {
            _showMarks = ShowMarksButton.IsChecked == true;
            UpdatePageBreakOverlay();
            TextEditor.Focus();
        }

        // Called from UpdatePageBreakOverlay after page lines are drawn
        private void DrawParagraphMarks()
        {
            if (!_showMarks)
                return;
            foreach (var para in EnumerateParagraphs(TextEditor.Document.Blocks))
            {
                try
                {
                    var rect = para.ContentEnd.GetCharacterRect(LogicalDirection.Backward);
                    if (rect.IsEmpty)
                        continue;
                    var mark = new TextBlock
                    {
                        Text = "¶",
                        Foreground = Brushes.SteelBlue,
                        FontSize = 14,
                        Opacity = 0.6
                    };
                    Canvas.SetLeft(mark, rect.Right + 1);
                    Canvas.SetTop(mark, rect.Top);
                    PageBreakOverlay.Children.Add(mark);
                }
                catch { }
            }
        }

        // Faintly render header/footer text and page numbers at each simulated
        // page boundary, plus solid lines for manual page breaks. This is an
        // approximation: the RichTextBox is a continuous surface, so the exact
        // layout is only guaranteed in Print Preview.
        private void DrawPageDecorations(double contentPerPage)
        {
            // Manual page breaks: a distinct solid blue line
            foreach (var para in EnumerateParagraphs(TextEditor.Document.Blocks))
            {
                if (para.Tag as string != "PageBreak")
                    continue;
                try
                {
                    var rect = para.ContentStart.GetCharacterRect(LogicalDirection.Forward);
                    if (rect.IsEmpty) continue;
                    PageBreakOverlay.Children.Add(new System.Windows.Shapes.Line
                    {
                        X1 = 0, X2 = _pageSize.WidthDip, Y1 = rect.Top, Y2 = rect.Top,
                        Stroke = Brushes.CornflowerBlue, StrokeThickness = 1.5
                    });
                }
                catch { }
            }

            if (_docSettings == null || !_docSettings.HasAnyContent)
                return;

            for (int page = 0; page < _totalPages; page++)
            {
                if (page == 0 && !_docSettings.ShowOnFirstPage)
                    continue;

                double pageTop = 80 + page * contentPerPage;
                double pageBottom = pageTop + contentPerPage;

                if (_docSettings.HeaderText.Length > 0)
                    AddOverlayText(_docSettings.HeaderText, 80, pageTop - 45, false);

                if (_docSettings.FooterText.Length > 0)
                    AddOverlayText(_docSettings.FooterText, 80, pageBottom + 12, false);

                if (_docSettings.ShowPageNumbers)
                {
                    var num = (page + 1).ToString();
                    switch (_docSettings.PageNumberPosition)
                    {
                        case PageNumberPosition.HeaderRight:
                            AddOverlayText(num, 0, pageTop - 45, true);
                            break;
                        case PageNumberPosition.FooterRight:
                            AddOverlayText(num, 0, pageBottom + 12, true);
                            break;
                        default:
                            AddOverlayText(num, 0, pageBottom + 12, false, center: true);
                            break;
                    }
                }
            }
        }

        private void AddOverlayText(string text, double left, double top, bool right, bool center = false)
        {
            var tb = new TextBlock
            {
                Text = text,
                Foreground = Brushes.Gray,
                FontSize = 13,
                Opacity = 0.75
            };
            if (right)
            {
                tb.Width = _pageSize.WidthDip - 80;
                tb.TextAlignment = TextAlignment.Right;
                Canvas.SetLeft(tb, 0);
            }
            else if (center)
            {
                tb.Width = _pageSize.WidthDip;
                tb.TextAlignment = TextAlignment.Center;
                Canvas.SetLeft(tb, 0);
            }
            else
            {
                Canvas.SetLeft(tb, left);
            }
            Canvas.SetTop(tb, top);
            PageBreakOverlay.Children.Add(tb);
        }

        private static IEnumerable<Paragraph> EnumerateParagraphs(BlockCollection blocks)
        {
            foreach (var block in blocks)
            {
                if (block is Paragraph p)
                    yield return p;
                else if (block is List list)
                    foreach (var li in list.ListItems)
                        foreach (var inner in EnumerateParagraphs(li.Blocks))
                            yield return inner;
                else if (block is Section sec)
                    foreach (var inner in EnumerateParagraphs(sec.Blocks))
                        yield return inner;
            }
        }

        #endregion

        #region Word count

        private void UpdateWordCount()
        {
            try
            {
                var text = new TextRange(TextEditor.Document.ContentStart, TextEditor.Document.ContentEnd).Text;
                int chars = text.Count(c => !char.IsControl(c));
                int words = text.Split(new[] { ' ', '\t', '\r', '\n', ' ' }, StringSplitOptions.RemoveEmptyEntries).Length;
                WordCountText.Text = string.Format(Loc.T("WordCount"), words, chars);
            }
            catch { }
        }

        #endregion

        #region Find / Replace / Hyperlink / Symbol / Page break

        private void MenuFind_Click(object sender, RoutedEventArgs e) => ShowFindReplace(false);
        private void MenuReplace_Click(object sender, RoutedEventArgs e) => ShowFindReplace(true);

        private void ShowFindReplace(bool showReplace)
        {
            if (_findDialog == null || !_findDialog.IsLoaded)
            {
                _findDialog = new FindReplaceDialog(TextEditor) { Owner = this };
                _findDialog.Closed += (s, e) => _findDialog = null;
            }
            _findDialog.SetReplaceVisible(showReplace);
            _findDialog.Show();
            _findDialog.Activate();
        }

        private void MenuHyperlink_Click(object sender, RoutedEventArgs e)
        {
            var selectedText = TextEditor.Selection.Text;
            var dlg = new HyperlinkDialog(selectedText) { Owner = this };
            if (dlg.ShowDialog() != true)
                return;

            var sel = TextEditor.Selection;
            if (!sel.IsEmpty)
                sel.Text = "";
            var link = new Hyperlink(new Run(dlg.DisplayText), sel.Start)
            {
                NavigateUri = TryUri(dlg.Url),
                ToolTip = dlg.Url
            };
            _hasUnsavedChanges = true;
            TextEditor.Focus();
        }

        private static Uri? TryUri(string url)
        {
            try { return new Uri(url, UriKind.RelativeOrAbsolute); }
            catch { return null; }
        }

        private void MenuSymbol_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new SymbolDialog { Owner = this };
            if (dlg.ShowDialog() == true && !string.IsNullOrEmpty(dlg.SelectedSymbol))
            {
                TextEditor.CaretPosition.InsertTextInRun(dlg.SelectedSymbol);
                _hasUnsavedChanges = true;
            }
            TextEditor.Focus();
        }

        private void MenuPageBreak_Click(object sender, RoutedEventArgs e)
        {
            var caret = TextEditor.CaretPosition;
            var para = caret.Paragraph;
            if (para == null)
                return;

            var breakPara = new Paragraph { Tag = "PageBreak" };
            para.SiblingBlocks?.InsertAfter(para, breakPara);
            var newCaret = breakPara.ContentEnd;
            TextEditor.CaretPosition = newCaret;
            _hasUnsavedChanges = true;
            UpdatePageBreakOverlay();
            TextEditor.Focus();
        }

        #endregion

        #region Export PDF

        private void MenuExportPdf_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var server = new System.Printing.LocalPrintServer();
                var pdfQueue = server.GetPrintQueues()
                    .FirstOrDefault(q => q.FullName.IndexOf("Microsoft Print to PDF", StringComparison.OrdinalIgnoreCase) >= 0);

                if (pdfQueue == null)
                {
                    MessageBox.Show(Loc.T("PdfPrinterMissing"), Loc.T("Error"), MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var printDlg = new PrintDialog { PrintQueue = pdfQueue };
                printDlg.PrintDocument(CreatePaginator(), "MiniWord");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, Loc.T("Error"), MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        #endregion

        #region Autosave / recovery

        private void MarkDraftDirty() => _draftDirty = true;

        private string DraftPath => Path.Combine(AutosaveDir, $"draft_{_windowIndex}.docx");
        private string DraftMarkerPath => Path.Combine(AutosaveDir, $"draft_{_windowIndex}.info");

        private void AutosaveTick()
        {
            if (!_draftDirty)
                return;
            _draftDirty = false;
            try
            {
                Directory.CreateDirectory(AutosaveDir);
                _docxService.SaveDocument(TextEditor.Document, DraftPath, _pageSize, _docSettings);
                File.WriteAllText(DraftMarkerPath, _currentFilePath ?? "");
            }
            catch { }
        }

        private void DeleteDraft()
        {
            try
            {
                if (File.Exists(DraftPath)) File.Delete(DraftPath);
                if (File.Exists(DraftMarkerPath)) File.Delete(DraftMarkerPath);
            }
            catch { }
        }

        // Offer to recover a draft left by a crash (primary window only)
        private void CheckForRecoverableDraft()
        {
            if (_windowIndex != 0)
                return;
            var path = DraftPath;
            if (!File.Exists(path))
                return;

            var result = MessageBox.Show(Loc.T("RecoverText"), Loc.T("RecoverTitle"),
                MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (result == MessageBoxResult.Yes)
            {
                try
                {
                    var doc = _docxService.LoadDocument(path);
                    TextEditor.Document = doc;
                    if (_docxService.LoadedPageSize != null)
                        ApplyPageSize(_docxService.LoadedPageSize);
                    _docSettings = _docxService.LoadedSettings ?? new DocumentSettings();
                    _hasUnsavedChanges = true;
                    UpdateWordCount();
                }
                catch { }
            }
            DeleteDraft();
        }

        #endregion
    }
}
