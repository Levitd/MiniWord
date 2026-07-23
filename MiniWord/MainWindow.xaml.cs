using System;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Xps;
using System.Windows.Xps.Packaging;
using Microsoft.Win32;
using MiniWord.Models;
using MiniWord.Services;

namespace MiniWord
{
    public partial class MainWindow : Window
    {
        private string? _currentFilePath;
        private bool _hasUnsavedChanges;
        private bool _updatingToolbar;
        private readonly DocxService _docxService = new();
        private readonly AppSettings _settings = AppSettings.Load();
        private PageSizeInfo _pageSize = PageSizeInfo.All[0];
        private Color _fontColor = Color.FromRgb(0xC0, 0x00, 0x00);
        private Color _highlightColor = Colors.Yellow;
        private DocumentSettings _docSettings = new();
        private int _totalPages = 1;

        // WPF font sizes are in DIP (1/96"), Word font sizes are in points (1/72")
        private static double PtToDip(double pt) => pt * 96.0 / 72.0;
        private static double DipToPt(double dip) => dip * 72.0 / 96.0;

        public MainWindow()
        {
            InitializeComponent();
            Loc.Lang = _settings.Language;

            InitializeToolbar();
            InitializeColorPalettes();
            ApplyEditorDefaults();
            ApplyPageSize(PageSizeInfo.ByName(_settings.PageSize));
            ApplyLocalization();
            SetupKeyBindings();

            TextEditor.TextChanged += TextEditor_TextChanged;
            TextEditor.SizeChanged += (s, e) => UpdatePageBreakOverlay();
            Closing += MainWindow_Closing;
        }

        #region Initialization

        private void InitializeToolbar()
        {
            _updatingToolbar = true;

            FontFamilyCombo.ItemsSource = Fonts.SystemFontFamilies
                .Select(f => f.Source)
                .OrderBy(s => s)
                .ToList();
            FontSizeCombo.ItemsSource = new[] { "8", "9", "10", "11", "12", "14", "16", "18", "20", "22", "24", "28", "36", "48", "72" };
            PageSizeCombo.ItemsSource = PageSizeInfo.All.Select(p => p.Name).ToList();

            _updatingToolbar = false;
        }

        private void InitializeColorPalettes()
        {
            Color[] fontColors =
            {
                Color.FromRgb(0x00, 0x00, 0x00), Color.FromRgb(0x44, 0x44, 0x44), Color.FromRgb(0x7F, 0x7F, 0x7F),
                Color.FromRgb(0xBF, 0xBF, 0xBF), Color.FromRgb(0xFF, 0xFF, 0xFF), Color.FromRgb(0xC0, 0x00, 0x00),
                Color.FromRgb(0xFF, 0x00, 0x00), Color.FromRgb(0xFF, 0xC0, 0x00), Color.FromRgb(0xFF, 0xFF, 0x00),
                Color.FromRgb(0x92, 0xD0, 0x50),
                Color.FromRgb(0x00, 0xB0, 0x50), Color.FromRgb(0x00, 0xB0, 0xF0), Color.FromRgb(0x00, 0x70, 0xC0),
                Color.FromRgb(0x00, 0x20, 0x60), Color.FromRgb(0x70, 0x30, 0xA0), Color.FromRgb(0xFF, 0x00, 0xFF),
                Color.FromRgb(0xFF, 0x99, 0xCC), Color.FromRgb(0x83, 0x3C, 0x00), Color.FromRgb(0x80, 0x80, 0x00),
                Color.FromRgb(0x00, 0x80, 0x80),
            };
            foreach (var c in fontColors)
                FontColorGrid.Children.Add(MakeSwatch(c, PickFontColor));

            // The 15 standard Word highlight colors
            Color[] highlightColors =
            {
                Colors.Yellow, Color.FromRgb(0x00, 0xFF, 0x00), Colors.Cyan, Colors.Magenta, Color.FromRgb(0x00, 0x00, 0xFF),
                Colors.Red, Color.FromRgb(0x00, 0x00, 0x8B), Color.FromRgb(0x00, 0x8B, 0x8B), Color.FromRgb(0x00, 0x64, 0x00), Color.FromRgb(0x8B, 0x00, 0x8B),
                Color.FromRgb(0x8B, 0x00, 0x00), Color.FromRgb(0x80, 0x80, 0x00), Color.FromRgb(0x80, 0x80, 0x80), Color.FromRgb(0xD3, 0xD3, 0xD3), Colors.Black,
            };
            foreach (var c in highlightColors)
                HighlightGrid.Children.Add(MakeSwatch(c, PickHighlightColor));
        }

        private Button MakeSwatch(Color color, Action<Color> onPick)
        {
            var button = new Button
            {
                Width = 18,
                Height = 18,
                Margin = new Thickness(1),
                Background = new SolidColorBrush(color),
                BorderBrush = Brushes.Gray,
                BorderThickness = new Thickness(1)
            };
            button.Click += (s, e) => onPick(color);
            return button;
        }

        private void ApplyEditorDefaults()
        {
            _updatingToolbar = true;

            TextEditor.FontFamily = new FontFamily(_settings.DefaultFontFamily);
            TextEditor.FontSize = PtToDip(_settings.DefaultFontSize);
            FontFamilyCombo.SelectedItem = _settings.DefaultFontFamily;
            FontSizeCombo.SelectedItem = _settings.DefaultFontSize.ToString("F0");

            _updatingToolbar = false;
        }

        private void ApplyPageSize(PageSizeInfo pageSize)
        {
            _pageSize = pageSize;
            TextEditor.Width = pageSize.WidthDip;
            TextEditor.MinHeight = pageSize.HeightDip;

            _updatingToolbar = true;
            PageSizeCombo.SelectedItem = pageSize.Name;
            _updatingToolbar = false;

            UpdatePageBreakOverlay();
        }

        // Approximate page boundaries in the continuous editing view:
        // each printed page holds (pageHeight - top/bottom margins) of content
        private void UpdatePageBreakOverlay()
        {
            PageBreakOverlay.Children.Clear();

            double contentPerPage = _pageSize.HeightDip - 160;
            double totalHeight = Math.Max(TextEditor.ActualHeight, _pageSize.HeightDip);
            _totalPages = Math.Max(1, (int)Math.Ceiling((totalHeight - 160) / contentPerPage));

            for (int i = 1; i < _totalPages; i++)
            {
                double y = 80 + i * contentPerPage;
                if (y > totalHeight - 40)
                    break;

                PageBreakOverlay.Children.Add(new System.Windows.Shapes.Line
                {
                    X1 = 0,
                    X2 = _pageSize.WidthDip,
                    Y1 = y,
                    Y2 = y,
                    Stroke = Brushes.LightSteelBlue,
                    StrokeThickness = 1,
                    StrokeDashArray = new DoubleCollection { 6, 4 }
                });
            }

            UpdatePageStatus();
        }

        private void UpdatePageStatus()
        {
            double contentPerPage = _pageSize.HeightDip - 160;
            var caretRect = TextEditor.CaretPosition.GetCharacterRect(LogicalDirection.Forward);
            int current = (int)((caretRect.Y - 80) / contentPerPage) + 1;
            current = Math.Max(1, Math.Min(_totalPages, current));
            PageInfoText.Text = string.Format(Loc.T("PageStatus"), current, _totalPages);
        }

        private void ApplyLocalization()
        {
            MenuFile.Header = Loc.T("File");
            MenuNew.Header = Loc.T("New");
            MenuOpen.Header = Loc.T("Open");
            MenuSave.Header = Loc.T("Save");
            MenuSaveAs.Header = Loc.T("SaveAs");
            MenuPreview.Header = Loc.T("Preview");
            MenuPrint.Header = Loc.T("Print");
            MenuExit.Header = Loc.T("Exit");
            MenuEdit.Header = Loc.T("Edit");
            MenuUndo.Header = Loc.T("Undo");
            MenuRedo.Header = Loc.T("Redo");
            MenuInsert.Header = Loc.T("Insert");
            MenuInsertImage.Header = Loc.T("Image");
            MenuHeaderFooter.Header = Loc.T("HeaderFooter");
            MenuTools.Header = Loc.T("Tools");
            MenuSettings.Header = Loc.T("Settings");

            OpenButton.ToolTip = Loc.T("TipOpen");
            SaveButton.ToolTip = Loc.T("TipSave");
            PrintButton.ToolTip = Loc.T("TipPrint");
            UndoButton.ToolTip = Loc.T("TipUndo");
            RedoButton.ToolTip = Loc.T("TipRedo");
            FontFamilyCombo.ToolTip = Loc.T("TipFont");
            FontSizeCombo.ToolTip = Loc.T("TipFontSize");
            BoldButton.ToolTip = Loc.T("TipBold");
            ItalicButton.ToolTip = Loc.T("TipItalic");
            UnderlineButton.ToolTip = Loc.T("TipUnderline");
            FontColorButton.ToolTip = Loc.T("TipFontColor");
            FontColorDropButton.ToolTip = Loc.T("TipFontColor");
            HighlightButton.ToolTip = Loc.T("TipHighlight");
            HighlightDropButton.ToolTip = Loc.T("TipHighlight");
            AutoColorButton.Content = Loc.T("Automatic");
            NoColorButton.Content = Loc.T("NoColor");
            AlignLeftButton.ToolTip = Loc.T("TipAlignLeft");
            AlignCenterButton.ToolTip = Loc.T("TipAlignCenter");
            AlignRightButton.ToolTip = Loc.T("TipAlignRight");
            AlignJustifyButton.ToolTip = Loc.T("TipAlignJustify");
            BulletsButton.ToolTip = Loc.T("TipBullets");
            NumberingButton.ToolTip = Loc.T("TipNumbering");
            ParagraphButton.ToolTip = Loc.T("TipParagraph");
            InsertImageButton.ToolTip = Loc.T("TipInsertImage");
            PageSizeCombo.ToolTip = Loc.T("TipPageSize");

            UpdateTitle();
            UpdatePageStatus();
        }

        private void UpdateTitle()
        {
            var name = string.IsNullOrEmpty(_currentFilePath)
                ? Loc.T("DefaultDocName")
                : Path.GetFileName(_currentFilePath);
            Title = $"{name} - MiniWord";
        }

        private void SetupKeyBindings()
        {
            AddKeyBinding(Key.N, (s, e) => MenuFileNew_Click(s, e));
            AddKeyBinding(Key.O, (s, e) => MenuFileOpen_Click(s, e));
            AddKeyBinding(Key.S, (s, e) => MenuFileSave_Click(s, e));
            AddKeyBinding(Key.P, (s, e) => MenuFilePrint_Click(s, e));
        }

        private void AddKeyBinding(Key key, ExecutedRoutedEventHandler handler)
        {
            var command = new RoutedCommand();
            command.InputGestures.Add(new KeyGesture(key, ModifierKeys.Control));
            CommandBindings.Add(new CommandBinding(command, handler));
        }

        #endregion

        private void TextEditor_TextChanged(object sender, TextChangedEventArgs e)
        {
            _hasUnsavedChanges = true;
        }

        private void UpdateToolbarState()
        {
            _updatingToolbar = true;

            var fontFamily = TextEditor.Selection.GetPropertyValue(TextBlock.FontFamilyProperty);
            if (fontFamily is FontFamily ff)
                FontFamilyCombo.SelectedItem = ff.Source;

            var fontSize = TextEditor.Selection.GetPropertyValue(TextBlock.FontSizeProperty);
            if (fontSize is double size)
                FontSizeCombo.SelectedItem = DipToPt(size).ToString("F0");

            var fontWeight = TextEditor.Selection.GetPropertyValue(TextBlock.FontWeightProperty);
            BoldButton.IsChecked = fontWeight is FontWeight fw && fw == FontWeights.Bold;

            var fontStyle = TextEditor.Selection.GetPropertyValue(TextBlock.FontStyleProperty);
            ItalicButton.IsChecked = fontStyle is FontStyle fs && fs == FontStyles.Italic;

            var decoration = TextEditor.Selection.GetPropertyValue(Inline.TextDecorationsProperty);
            UnderlineButton.IsChecked = decoration is TextDecorationCollection tdc && tdc.Count > 0;

            var alignment = TextEditor.Selection.Start.Paragraph?.TextAlignment;
            AlignLeftButton.IsChecked = alignment == TextAlignment.Left;
            AlignCenterButton.IsChecked = alignment == TextAlignment.Center;
            AlignRightButton.IsChecked = alignment == TextAlignment.Right;
            AlignJustifyButton.IsChecked = alignment == TextAlignment.Justify;

            _updatingToolbar = false;
        }

        #region Menu - File

        private void MenuFileNew_Click(object sender, RoutedEventArgs e)
        {
            if (_hasUnsavedChanges && !ConfirmUnsavedChanges())
                return;

            TextEditor.Document = new FlowDocument(new Paragraph());
            ApplyEditorDefaults();
            ApplyPageSize(PageSizeInfo.ByName(_settings.PageSize));
            _docSettings = new DocumentSettings();
            _currentFilePath = null;
            _hasUnsavedChanges = false;
            UpdateTitle();
        }

        private void MenuFileOpen_Click(object sender, RoutedEventArgs e)
        {
            if (_hasUnsavedChanges && !ConfirmUnsavedChanges())
                return;

            var dlg = new OpenFileDialog
            {
                Filter = Loc.T("FilterDocx"),
                DefaultExt = ".docx"
            };

            if (dlg.ShowDialog() != true)
                return;

            try
            {
                var doc = _docxService.LoadDocument(dlg.FileName);
                TextEditor.Document = doc;
                if (_docxService.LoadedPageSize != null)
                    ApplyPageSize(_docxService.LoadedPageSize);
                _docSettings = _docxService.LoadedSettings ?? new DocumentSettings();
                _currentFilePath = dlg.FileName;
                _hasUnsavedChanges = false;
                UpdateTitle();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"{Loc.T("ErrorOpen")}: {ex.Message}", Loc.T("Error"), MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void MenuFileSave_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(_currentFilePath))
            {
                MenuFileSaveAs_Click(sender, e);
                return;
            }

            SaveDocumentToFile(_currentFilePath);
        }

        private void MenuFileSaveAs_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new SaveFileDialog
            {
                Filter = Loc.T("FilterDocxSave"),
                DefaultExt = ".docx",
                FileName = string.IsNullOrEmpty(_currentFilePath)
                    ? Loc.T("DefaultDocName")
                    : Path.GetFileNameWithoutExtension(_currentFilePath)
            };

            if (dlg.ShowDialog() != true)
                return;

            SaveDocumentToFile(dlg.FileName);
            _currentFilePath = dlg.FileName;
            UpdateTitle();
        }

        private void SaveDocumentToFile(string filePath)
        {
            try
            {
                _docxService.SaveDocument(TextEditor.Document, filePath, _pageSize, _docSettings);
                _hasUnsavedChanges = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"{Loc.T("ErrorSave")}: {ex.Message}", Loc.T("Error"), MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // Print/preview work on a serialized copy so pagination never
        // touches the document that is live inside the RichTextBox
        private FlowDocument ClonePrintDocument()
        {
            var range = new TextRange(TextEditor.Document.ContentStart, TextEditor.Document.ContentEnd);
            using var ms = new MemoryStream();
            range.Save(ms, DataFormats.XamlPackage);
            ms.Position = 0;

            var clone = new FlowDocument();
            new TextRange(clone.ContentStart, clone.ContentEnd).Load(ms, DataFormats.XamlPackage);
            clone.FontFamily = TextEditor.FontFamily;
            clone.FontSize = TextEditor.FontSize;
            return clone;
        }

        private DocumentPaginator CreatePaginator() =>
            new(ClonePrintDocument(), new Size(_pageSize.WidthDip, _pageSize.HeightDip), _docSettings);

        private void MenuFilePrint_Click(object sender, RoutedEventArgs e)
        {
            var printDlg = new PrintDialog();
            if (printDlg.ShowDialog() == true)
                printDlg.PrintDocument(CreatePaginator(), "MiniWord");
        }

        private void MenuFilePreview_Click(object sender, RoutedEventArgs e)
        {
            var xpsPath = Path.Combine(Path.GetTempPath(), $"MiniWord_{Guid.NewGuid():N}.xps");
            try
            {
                using var xpsDoc = new XpsDocument(xpsPath, FileAccess.ReadWrite);
                XpsDocument.CreateXpsDocumentWriter(xpsDoc).Write(CreatePaginator());

                var win = new PreviewWindow { Owner = this };
                win.Viewer.Document = xpsDoc.GetFixedDocumentSequence();
                win.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, Loc.T("Error"), MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                try { File.Delete(xpsPath); } catch { }
            }
        }

        private void MenuFileExit_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        #endregion

        #region Menu - Edit

        private void MenuEditUndo_Click(object sender, RoutedEventArgs e)
        {
            if (TextEditor.CanUndo)
                TextEditor.Undo();
        }

        private void MenuEditRedo_Click(object sender, RoutedEventArgs e)
        {
            if (TextEditor.CanRedo)
                TextEditor.Redo();
        }

        #endregion

        #region Menu - Insert

        private void MenuInsertImage_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog { Filter = Loc.T("FilterImages") };
            if (dlg.ShowDialog() != true)
                return;

            try
            {
                var bmp = new BitmapImage();
                bmp.BeginInit();
                bmp.CacheOption = BitmapCacheOption.OnLoad;
                bmp.UriSource = new Uri(dlg.FileName);
                bmp.EndInit();
                bmp.Freeze();

                // Fit into text area width (page minus margins)
                double maxWidth = _pageSize.WidthDip - TextEditor.Padding.Left - TextEditor.Padding.Right;
                double w = bmp.Width;
                double h = bmp.Height;
                if (w > maxWidth)
                {
                    h *= maxWidth / w;
                    w = maxWidth;
                }

                var image = new Image { Source = bmp, Width = w, Height = h, Stretch = Stretch.Uniform };
                _ = new InlineUIContainer(image, TextEditor.CaretPosition);
                _hasUnsavedChanges = true;
                TextEditor.Focus();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"{Loc.T("ErrorOpen")}: {ex.Message}", Loc.T("Error"), MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void MenuHeaderFooter_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new HeaderFooterDialog(_docSettings) { Owner = this };
            if (dlg.ShowDialog() == true)
                _hasUnsavedChanges = true;
        }

        #endregion

        #region Menu - Tools

        private void MenuSettings_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new SettingsDialog(_settings) { Owner = this };
            if (dlg.ShowDialog() == true)
            {
                Loc.Lang = _settings.Language;
                ApplyEditorDefaults();
                ApplyLocalization();
            }
        }

        #endregion

        #region Toolbar - Font

        private void FontFamilyCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_updatingToolbar)
                return;

            if (FontFamilyCombo.SelectedItem is string fontName)
            {
                TextEditor.Selection.ApplyPropertyValue(TextBlock.FontFamilyProperty, new FontFamily(fontName));
                TextEditor.Focus();
            }
        }

        private void FontSizeCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_updatingToolbar)
                return;

            if (FontSizeCombo.SelectedItem is string sizeStr && double.TryParse(sizeStr, out var sizePt))
            {
                TextEditor.Selection.ApplyPropertyValue(TextBlock.FontSizeProperty, PtToDip(sizePt));
                TextEditor.Focus();
            }
        }

        #endregion

        #region Toolbar - Character Formatting

        private void BoldButton_Click(object sender, RoutedEventArgs e)
        {
            var weight = TextEditor.Selection.GetPropertyValue(TextBlock.FontWeightProperty);
            bool isBold = weight is FontWeight fw && fw == FontWeights.Bold;
            TextEditor.Selection.ApplyPropertyValue(TextBlock.FontWeightProperty,
                isBold ? FontWeights.Normal : FontWeights.Bold);
            UpdateToolbarState();
            TextEditor.Focus();
        }

        private void ItalicButton_Click(object sender, RoutedEventArgs e)
        {
            var style = TextEditor.Selection.GetPropertyValue(TextBlock.FontStyleProperty);
            bool isItalic = style is FontStyle fs && fs == FontStyles.Italic;
            TextEditor.Selection.ApplyPropertyValue(TextBlock.FontStyleProperty,
                isItalic ? FontStyles.Normal : FontStyles.Italic);
            UpdateToolbarState();
            TextEditor.Focus();
        }

        private void UnderlineButton_Click(object sender, RoutedEventArgs e)
        {
            var decoration = TextEditor.Selection.GetPropertyValue(Inline.TextDecorationsProperty);
            bool isUnderlined = decoration is TextDecorationCollection tdc && tdc.Count > 0;
            TextEditor.Selection.ApplyPropertyValue(Inline.TextDecorationsProperty,
                isUnderlined ? null : TextDecorations.Underline);
            UpdateToolbarState();
            TextEditor.Focus();
        }

        #endregion

        #region Toolbar - Colors

        private void FontColorButton_Click(object sender, RoutedEventArgs e)
        {
            ApplySelectionBrush(TextElement.ForegroundProperty, new SolidColorBrush(_fontColor));
        }

        private void FontColorDrop_Click(object sender, RoutedEventArgs e)
        {
            FontColorPopup.IsOpen = true;
        }

        private void PickFontColor(Color color)
        {
            _fontColor = color;
            FontColorBar.Fill = new SolidColorBrush(color);
            FontColorPopup.IsOpen = false;
            ApplySelectionBrush(TextElement.ForegroundProperty, new SolidColorBrush(color));
        }

        private void AutoColor_Click(object sender, RoutedEventArgs e)
        {
            _fontColor = Colors.Black;
            FontColorBar.Fill = Brushes.Black;
            FontColorPopup.IsOpen = false;
            ApplySelectionBrush(TextElement.ForegroundProperty, Brushes.Black);
        }

        private void HighlightButton_Click(object sender, RoutedEventArgs e)
        {
            ApplySelectionBrush(TextElement.BackgroundProperty, new SolidColorBrush(_highlightColor));
        }

        private void HighlightDrop_Click(object sender, RoutedEventArgs e)
        {
            HighlightPopup.IsOpen = true;
        }

        private void PickHighlightColor(Color color)
        {
            _highlightColor = color;
            HighlightBar.Fill = new SolidColorBrush(color);
            HighlightPopup.IsOpen = false;
            ApplySelectionBrush(TextElement.BackgroundProperty, new SolidColorBrush(color));
        }

        private void NoHighlight_Click(object sender, RoutedEventArgs e)
        {
            HighlightPopup.IsOpen = false;
            ApplySelectionBrush(TextElement.BackgroundProperty, null);
        }

        private void ApplySelectionBrush(DependencyProperty property, Brush? brush)
        {
            TextEditor.Selection.ApplyPropertyValue(property, brush);
            _hasUnsavedChanges = true;
            TextEditor.Focus();
        }

        #endregion

        #region Toolbar - Paragraph

        private void AlignLeft_Click(object sender, RoutedEventArgs e) => ApplyParagraphAlignment(TextAlignment.Left);
        private void AlignCenter_Click(object sender, RoutedEventArgs e) => ApplyParagraphAlignment(TextAlignment.Center);
        private void AlignRight_Click(object sender, RoutedEventArgs e) => ApplyParagraphAlignment(TextAlignment.Right);
        private void AlignJustify_Click(object sender, RoutedEventArgs e) => ApplyParagraphAlignment(TextAlignment.Justify);

        private void ApplyParagraphAlignment(TextAlignment alignment)
        {
            var p = TextEditor.Selection.Start.Paragraph;
            if (p != null)
            {
                p.TextAlignment = alignment;
                _hasUnsavedChanges = true;
            }
            UpdateToolbarState();
            TextEditor.Focus();
        }

        private void Bullets_Click(object sender, RoutedEventArgs e)
        {
            EditingCommands.ToggleBullets.Execute(null, TextEditor);
            _hasUnsavedChanges = true;
            TextEditor.Focus();
        }

        private void Numbering_Click(object sender, RoutedEventArgs e)
        {
            EditingCommands.ToggleNumbering.Execute(null, TextEditor);
            _hasUnsavedChanges = true;
            TextEditor.Focus();
        }

        private void ParagraphDialog_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new ParagraphDialog(TextEditor) { Owner = this };
            if (dlg.ShowDialog() == true)
                _hasUnsavedChanges = true;
        }

        private void PageSizeCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_updatingToolbar)
                return;

            if (PageSizeCombo.SelectedItem is string name)
            {
                ApplyPageSize(PageSizeInfo.ByName(name));
                _settings.PageSize = name;
                _settings.Save();
                _hasUnsavedChanges = true;
            }
        }

        #endregion

        #region Window Lifecycle

        private void MainWindow_Closing(object? sender, System.ComponentModel.CancelEventArgs e)
        {
            if (_hasUnsavedChanges && !ConfirmUnsavedChanges())
                e.Cancel = true;
        }

        private bool ConfirmUnsavedChanges()
        {
            var result = MessageBox.Show(
                Loc.T("UnsavedText"),
                Loc.T("UnsavedTitle"),
                MessageBoxButton.YesNoCancel,
                MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                MenuFileSave_Click(this, new RoutedEventArgs());
                return !_hasUnsavedChanges;
            }

            return result == MessageBoxResult.No;
        }

        #endregion

        private void TextEditor_SelectionChanged(object sender, RoutedEventArgs e)
        {
            UpdateToolbarState();
            UpdatePageStatus();
        }
    }
}
