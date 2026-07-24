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
        private readonly AppSettings _settings = AppSettings.Current;
        private PageSizeInfo _pageSize = PageSizeInfo.All[0];
        private Color _fontColor = Color.FromRgb(0xC0, 0x00, 0x00);
        private Color _highlightColor = Colors.Yellow;
        private DocumentSettings _docSettings = new();
        private int _totalPages = 1;
        private readonly string? _initialFile;

        // Editable page view: extra top-margins that push each page's first block
        // onto a fresh sheet, creating a real gap. Stripped before save/print.
        private bool _paginating;
        private readonly System.Collections.Generic.List<(System.Windows.Documents.Block block, Thickness orig)> _pageGapMargins = new();
        private const double PageGap = 24;

        // Shared across all windows in the process
        private static bool _updateCheckedThisProcess;
        private static int _windowsCreated;
        private readonly int _windowIndex;

        // WPF font sizes are in DIP (1/96"), Word font sizes are in points (1/72")
        private static double PtToDip(double pt) => pt * 96.0 / 72.0;
        private static double DipToPt(double dip) => dip * 72.0 / 96.0;

        public MainWindow() : this(null) { }

        public MainWindow(string? filePath)
        {
            _initialFile = filePath;
            _windowIndex = _windowsCreated++;

            InitializeComponent();
            Loc.Lang = _settings.Language;
            RestoreWindowBounds();

            InitializeToolbar();
            InitializeColorPalettes();
            ApplyEditorDefaults();
            ApplyPageSize(PageSizeInfo.ByName(_settings.PageSize));
            ApplyLocalization();
            RebuildRecentMenu();
            SetupKeyBindings();

            TextEditor.TextChanged += TextEditor_TextChanged;
            TextEditor.SizeChanged += (s, e) => UpdatePageBreakOverlay();
            Closing += MainWindow_Closing;
            Loaded += MainWindow_Loaded;

            InitializeFeatures();
        }

        private async void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(_initialFile))
            {
                LoadFile(_initialFile);
            }
            else
            {
                CheckForRecoverableDraft();
                if (_windowIndex == 0 && _settings.RecentFiles.Count > 0 && !_hasUnsavedChanges)
                    ShowStartScreen();
            }

            // Update handling runs once for the whole process, from the first window
            if (_windowIndex == 0 && !_updateCheckedThisProcess)
            {
                _updateCheckedThisProcess = true;

                if (TryPromptPendingUpdate())
                    return;

                var info = await UpdateService.CheckForUpdateAsync();
                if (info != null)
                    new UpdateDialog(info) { Owner = this }.ShowDialog();
            }
        }

        private void ShowStartScreen()
        {
            var dlg = new StartDialog(_settings.RecentFiles) { Owner = this };
            if (dlg.ShowDialog() == true && !string.IsNullOrEmpty(dlg.SelectedFile))
                LoadFile(dlg.SelectedFile);
        }

        #region Initialization

        private void RestoreWindowBounds()
        {
            if (_settings.WindowWidth > 300 && _settings.WindowHeight > 200)
            {
                Width = _settings.WindowWidth;
                Height = _settings.WindowHeight;
            }

            // Restore position only if it is still on a connected screen
            if (_settings.WindowLeft.HasValue && _settings.WindowTop.HasValue
                && _settings.WindowLeft.Value > SystemParameters.VirtualScreenLeft - 50
                && _settings.WindowTop.Value > SystemParameters.VirtualScreenTop - 50
                && _settings.WindowLeft.Value < SystemParameters.VirtualScreenLeft + SystemParameters.VirtualScreenWidth - 100
                && _settings.WindowTop.Value < SystemParameters.VirtualScreenTop + SystemParameters.VirtualScreenHeight - 100)
            {
                WindowStartupLocation = WindowStartupLocation.Manual;
                Left = _settings.WindowLeft.Value;
                Top = _settings.WindowTop.Value;
            }

            if (_settings.WindowMaximized && _windowIndex == 0)
                WindowState = WindowState.Maximized;

            // Cascade additional windows so they don't stack exactly
            if (_windowIndex > 0)
            {
                WindowStartupLocation = WindowStartupLocation.Manual;
                WindowState = WindowState.Normal;
                double offset = 30 * _windowIndex;
                Left = (_settings.WindowLeft ?? 100) + offset;
                Top = (_settings.WindowTop ?? 100) + offset;
            }
        }

        private void SaveWindowBounds()
        {
            // Only the primary window owns the saved bounds; cascaded windows
            // would otherwise overwrite them with offset positions.
            if (_windowIndex != 0)
                return;

            _settings.WindowMaximized = WindowState == WindowState.Maximized;

            // For a maximized window store its restored (normal) size
            var bounds = WindowState == WindowState.Normal
                ? new Rect(Left, Top, Width, Height)
                : RestoreBounds;
            if (bounds.Width > 300 && bounds.Height > 200)
            {
                _settings.WindowWidth = bounds.Width;
                _settings.WindowHeight = bounds.Height;
                _settings.WindowLeft = bounds.Left;
                _settings.WindowTop = bounds.Top;
            }

            _settings.Save();
        }

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

        private void UpdatePageBreakOverlay()
        {
            if (_paginating)
                return;
            _paginating = true;
            try { RepaginateSheets(); }
            finally { _paginating = false; }
        }

        // Editable "page view": lay the document out as separate white sheets with
        // a gray gap between them. The RichTextBox is one continuous surface, so a
        // real gap is produced by giving the first block of each page an extra top
        // margin, snapping automatic breaks to paragraph boundaries. These synthetic
        // margins are stripped before the document is saved or printed (WithGapsRemoved).
        private void RepaginateSheets()
        {
            PageBreakOverlay.Children.Clear();
            PageSheetLayer.Children.Clear();

            double pw = _pageSize.WidthDip;
            double ph = _pageSize.HeightDip;
            double cpp = ph - 160;          // content height per sheet (80 top + 80 bottom)
            double extra = 160 + PageGap;   // bottom margin + gap + top margin between sheets

            // 1. Remove previous gap margins so we measure the continuous flow.
            foreach (var (blk, orig) in _pageGapMargins)
                blk.Margin = orig;
            _pageGapMargins.Clear();
            TextEditor.UpdateLayout();

            // 2. Decide which top-level blocks start a new page (by continuous position).
            var starts = new System.Collections.Generic.List<(Block block, int page)>();
            int prevPage = 0, maxPage = 0;
            foreach (var block in TextEditor.Document.Blocks)
            {
                double y;
                try { y = block.ContentStart.GetCharacterRect(LogicalDirection.Forward).Top; }
                catch { continue; }
                if (double.IsNaN(y) || double.IsInfinity(y))
                    continue;

                int page = Math.Max(0, (int)((y - 80 + 2) / cpp));
                if ((block as Paragraph)?.Tag as string == "PageBreak" && page <= prevPage)
                    page = prevPage + 1;      // a manual break always advances a page
                if (page > prevPage)
                {
                    starts.Add((block, page));
                    prevPage = page;
                }
                maxPage = Math.Max(maxPage, page);
            }
            _totalPages = maxPage + 1;

            // 3. Apply gap margins (each page jump adds one sheet's worth of spacing).
            int fromPage = 0;
            foreach (var (block, page) in starts)
            {
                int jump = page - fromPage;
                fromPage = page;
                var orig = block.Margin;
                _pageGapMargins.Add((block, orig));
                block.Margin = new Thickness(orig.Left, orig.Top + extra * jump, orig.Right, orig.Bottom);
            }

            double stackHeight = _totalPages * ph + (_totalPages - 1) * PageGap;
            TextEditor.MinHeight = stackHeight;
            TextEditor.UpdateLayout();

            // 4. Draw the white sheets on the layer behind the transparent editor.
            for (int k = 0; k < _totalPages; k++)
            {
                var sheet = new System.Windows.Shapes.Rectangle
                {
                    Width = pw,
                    Height = ph,
                    Fill = Brushes.White,
                    Effect = new System.Windows.Media.Effects.DropShadowEffect
                    {
                        Color = Colors.Black,
                        Opacity = 0.35,
                        BlurRadius = 12,
                        ShadowDepth = 3
                    }
                };
                Canvas.SetLeft(sheet, 0);
                Canvas.SetTop(sheet, k * (ph + PageGap));
                PageSheetLayer.Children.Add(sheet);
            }

            DrawParagraphMarks();
            DrawPageDecorations();
            UpdatePageStatus();
        }

        // Remove the synthetic page-gap margins, run an action against the plain
        // continuous document (save/print/export), then restore the page view so
        // the gaps never leak into files.
        private void WithGapsRemoved(Action action)
        {
            bool had = _pageGapMargins.Count > 0;
            foreach (var (blk, orig) in _pageGapMargins)
                blk.Margin = orig;
            _pageGapMargins.Clear();
            try { action(); }
            finally { if (had) UpdatePageBreakOverlay(); }
        }

        private void UpdatePageStatus()
        {
            double step = _pageSize.HeightDip + PageGap;
            double caretY;
            try { caretY = TextEditor.CaretPosition.GetCharacterRect(LogicalDirection.Forward).Y; }
            catch { caretY = 0; }
            int current = (int)(caretY / step) + 1;
            current = Math.Max(1, Math.Min(_totalPages, current));
            PageInfoText.Text = string.Format(Loc.T("PageStatus"), current, _totalPages);
        }

        private void ApplyLocalization()
        {
            MenuFile.Header = Loc.T("File");
            MenuNew.Header = Loc.T("New");
            MenuNewWindow.Header = Loc.T("NewWindow");
            MenuRecent.Header = Loc.T("RecentFiles");
            MenuOpen.Header = Loc.T("Open");
            MenuSave.Header = Loc.T("Save");
            MenuSaveAs.Header = Loc.T("SaveAs");
            MenuPreview.Header = Loc.T("Preview");
            MenuPrint.Header = Loc.T("Print");
            MenuExit.Header = Loc.T("Exit");
            MenuEdit.Header = Loc.T("Edit");
            MenuUndo.Header = Loc.T("Undo");
            MenuRedo.Header = Loc.T("Redo");
            MenuCut.Header = Loc.T("Cut");
            MenuCopy.Header = Loc.T("Copy");
            MenuPaste.Header = Loc.T("Paste");
            MenuPasteText.Header = Loc.T("PasteText");
            MenuSelectAll.Header = Loc.T("SelectAll");
            MenuFind.Header = Loc.T("Find");
            MenuReplace.Header = Loc.T("Replace");
            MenuInsert.Header = Loc.T("Insert");
            MenuInsertImage.Header = Loc.T("Image");
            MenuInsertHyperlink.Header = Loc.T("Hyperlink");
            MenuInsertSymbol.Header = Loc.T("Symbol");
            MenuInsertPageBreak.Header = Loc.T("PageBreak");
            MenuHeaderFooter.Header = Loc.T("HeaderFooter");
            MenuExportPdf.Header = Loc.T("ExportPdf");
            MenuExportHtml.Header = Loc.T("ExportHtml");
            MenuTools.Header = Loc.T("Tools");
            MenuSettings.Header = Loc.T("Settings");
            MenuHelp.Header = Loc.T("Help");
            MenuCheckUpdates.Header = Loc.T("CheckUpdates");
            MenuAbout.Header = Loc.T("About");

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
            StrikeButton.ToolTip = Loc.T("TipStrikethrough");
            SuperscriptButton.ToolTip = Loc.T("TipSuperscript");
            SubscriptButton.ToolTip = Loc.T("TipSubscript");
            ChangeCaseButton.ToolTip = Loc.T("TipChangeCase");
            CutButton.ToolTip = Loc.T("TipCut");
            CopyButton.ToolTip = Loc.T("TipCopy");
            PasteButton.ToolTip = Loc.T("TipPaste");
            FormatPainterButton.ToolTip = Loc.T("TipFormatPainter");
            LineSpacingButton.ToolTip = Loc.T("TipLineSpacing");
            ShowMarksButton.ToolTip = Loc.T("TipShowMarks");
            CaseSentenceItem.Header = Loc.T("CaseSentence");
            CaseLowerItem.Header = Loc.T("CaseLower");
            CaseUpperItem.Header = Loc.T("CaseUpper");
            CaseTitleItem.Header = Loc.T("CaseTitle");
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
            AddKeyBinding(Key.N, ModifierKeys.Control, (s, e) => MenuFileNew_Click(s, e));
            AddKeyBinding(Key.N, ModifierKeys.Control | ModifierKeys.Shift, (s, e) => MenuNewWindow_Click(s, e));
            AddKeyBinding(Key.O, ModifierKeys.Control, (s, e) => MenuFileOpen_Click(s, e));
            AddKeyBinding(Key.S, ModifierKeys.Control, (s, e) => MenuFileSave_Click(s, e));
            AddKeyBinding(Key.P, ModifierKeys.Control, (s, e) => MenuFilePrint_Click(s, e));
            AddKeyBinding(Key.V, ModifierKeys.Control | ModifierKeys.Shift, (s, e) => MenuPasteText_Click(s, e));
            AddKeyBinding(Key.F, ModifierKeys.Control, (s, e) => MenuFind_Click(s, e));
            AddKeyBinding(Key.H, ModifierKeys.Control, (s, e) => MenuReplace_Click(s, e));
            AddKeyBinding(Key.K, ModifierKeys.Control, (s, e) => MenuHyperlink_Click(s, e));
            AddKeyBinding(Key.Enter, ModifierKeys.Control, (s, e) => MenuPageBreak_Click(s, e));
        }

        private void AddKeyBinding(Key key, ModifierKeys modifiers, ExecutedRoutedEventHandler handler)
        {
            var command = new RoutedCommand();
            command.InputGestures.Add(new KeyGesture(key, modifiers));
            CommandBindings.Add(new CommandBinding(command, handler));
        }

        #endregion

        private void TextEditor_TextChanged(object sender, TextChangedEventArgs e)
        {
            _hasUnsavedChanges = true;
            UpdateWordCount();
            MarkDraftDirty();
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

            var decoration = TextEditor.Selection.GetPropertyValue(Inline.TextDecorationsProperty) as TextDecorationCollection;
            UnderlineButton.IsChecked = HasDecoration(decoration, TextDecorationLocation.Underline);
            StrikeButton.IsChecked = HasDecoration(decoration, TextDecorationLocation.Strikethrough);

            var baseline = TextEditor.Selection.GetPropertyValue(Inline.BaselineAlignmentProperty);
            SuperscriptButton.IsChecked = baseline is BaselineAlignment ba1 && ba1 == BaselineAlignment.Superscript;
            SubscriptButton.IsChecked = baseline is BaselineAlignment ba2 && ba2 == BaselineAlignment.Subscript;

            var alignment = TextEditor.Selection.Start.Paragraph?.TextAlignment;
            AlignLeftButton.IsChecked = alignment == TextAlignment.Left;
            AlignCenterButton.IsChecked = alignment == TextAlignment.Center;
            AlignRightButton.IsChecked = alignment == TextAlignment.Right;
            AlignJustifyButton.IsChecked = alignment == TextAlignment.Justify;

            _updatingToolbar = false;
        }

        private static bool HasDecoration(TextDecorationCollection? decorations, TextDecorationLocation location)
        {
            if (decorations == null)
                return false;
            foreach (var d in decorations)
                if (d.Location == location)
                    return true;
            return false;
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
            var dlg = new OpenFileDialog
            {
                Filter = Loc.T("FilterDocx"),
                DefaultExt = ".docx",
                Multiselect = true
            };

            if (dlg.ShowDialog() != true)
                return;

            var files = dlg.FileNames;

            // First file goes into this window (if it can accept it),
            // the rest each open in a new window.
            int startIndex = 0;
            if (CanReplaceCurrentDocument())
            {
                LoadFile(files[0]);
                startIndex = 1;
            }

            for (int i = startIndex; i < files.Length; i++)
                OpenInNewWindow(files[i]);
        }

        private bool CanReplaceCurrentDocument()
        {
            if (!_hasUnsavedChanges)
                return true;
            return ConfirmUnsavedChanges();
        }

        /// <summary>Picks the reader/writer for a path by its extension (.docx default).</summary>
        private IDocumentFormat GetFormatService(string path) =>
            Path.GetExtension(path).ToLowerInvariant() switch
            {
                ".rtf" => new RtfService(),
                ".txt" => new TxtService(),
                ".odt" => new OdtService(),
                _ => _docxService,
            };

        /// <summary>Loads a document (docx/rtf/odt/txt) into this window.</summary>
        public void LoadFile(string path)
        {
            try
            {
                var svc = GetFormatService(path);
                var doc = svc.LoadDocument(path);
                TextEditor.Document = doc;
                if (svc.LoadedPageSize != null)
                    ApplyPageSize(svc.LoadedPageSize);
                _docSettings = svc.LoadedSettings ?? new DocumentSettings();
                _currentFilePath = path;
                _hasUnsavedChanges = false;
                UpdateTitle();

                _settings.AddRecentFile(path);
                RebuildRecentMenu();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"{Loc.T("ErrorOpen")}: {ex.Message}", Loc.T("Error"), MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void OpenInNewWindow(string? path)
        {
            var win = new MainWindow(path);
            win.Show();
        }

        private void MenuNewWindow_Click(object sender, RoutedEventArgs e) => OpenInNewWindow(null);

        private void RebuildRecentMenu()
        {
            MenuRecent.Items.Clear();
            var recent = _settings.RecentFiles;
            MenuRecent.IsEnabled = recent.Count > 0;

            for (int i = 0; i < recent.Count; i++)
            {
                var path = recent[i];
                var item = new MenuItem { Header = $"_{i + 1}  {Path.GetFileName(path)}", ToolTip = path };
                item.Click += (s, _) => OpenRecent(path);
                MenuRecent.Items.Add(item);
            }
        }

        private void OpenRecent(string path)
        {
            if (!File.Exists(path))
            {
                MessageBox.Show($"{Loc.T("ErrorOpen")}: {path}", Loc.T("Error"), MessageBoxButton.OK, MessageBoxImage.Warning);
                _settings.RecentFiles.Remove(path);
                _settings.Save();
                RebuildRecentMenu();
                return;
            }

            if (CanReplaceCurrentDocument())
                LoadFile(path);
            else
                OpenInNewWindow(path);
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
                WithGapsRemoved(() =>
                    GetFormatService(filePath).SaveDocument(TextEditor.Document, filePath, _pageSize, _docSettings));
                _hasUnsavedChanges = false;
                _draftDirty = false;
                DeleteDraft();
                _settings.AddRecentFile(filePath);
                RebuildRecentMenu();
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
            var clone = new FlowDocument();
            // Serialize the continuous document (no page-gap margins) so printed
            // output is not doubly spaced by the on-screen sheet gaps.
            WithGapsRemoved(() =>
            {
                var range = new TextRange(TextEditor.Document.ContentStart, TextEditor.Document.ContentEnd);
                using var ms = new MemoryStream();
                range.Save(ms, DataFormats.XamlPackage);
                ms.Position = 0;
                new TextRange(clone.ContentStart, clone.ContentEnd).Load(ms, DataFormats.XamlPackage);
            });
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
            {
                _hasUnsavedChanges = true;
                UpdatePageBreakOverlay();
            }
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

        #region Menu - Help

        private async void MenuCheckUpdates_Click(object sender, RoutedEventArgs e)
        {
            MenuCheckUpdates.IsEnabled = false;
            try
            {
                var info = await UpdateService.CheckForUpdateAsync();
                if (info != null)
                {
                    new UpdateDialog(info) { Owner = this }.ShowDialog();
                }
                else
                {
                    MessageBox.Show(
                        string.Format(Loc.T("UpToDateText"), UpdateService.CurrentVersion.ToString(3)),
                        Loc.T("UpToDateTitle"), MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            finally
            {
                MenuCheckUpdates.IsEnabled = true;
            }
        }

        private void MenuAbout_Click(object sender, RoutedEventArgs e)
        {
            new AboutDialog { Owner = this }.ShowDialog();
        }

        /// <summary>
        /// If a background download finished earlier, offer to install it now.
        /// Returns true if the app is shutting down to install.
        /// </summary>
        private bool TryPromptPendingUpdate()
        {
            var path = _settings.PendingUpdatePath;
            if (string.IsNullOrEmpty(path) || !File.Exists(path))
            {
                ClearPendingUpdate();
                return false;
            }

            var result = MessageBox.Show(
                string.Format(Loc.T("PendingUpdateText"), _settings.PendingUpdateVersion),
                Loc.T("UpdateTitle"), MessageBoxButton.YesNo, MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                UpdateService.LaunchInstaller(path);
                Application.Current.Shutdown();
                return true;
            }

            return false;
        }

        private void ClearPendingUpdate()
        {
            if (_settings.PendingUpdatePath == null && _settings.PendingUpdateVersion == null)
                return;
            _settings.PendingUpdatePath = null;
            _settings.PendingUpdateVersion = null;
            _settings.Save();
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
            ToggleTextDecoration(TextDecorations.Underline[0]);
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
            // Edit paragraph spacing against the continuous document so the dialog
            // never sees (or bakes in) the synthetic page-gap margins.
            WithGapsRemoved(() =>
            {
                var dlg = new ParagraphDialog(TextEditor) { Owner = this };
                if (dlg.ShowDialog() == true)
                    _hasUnsavedChanges = true;
            });
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
            {
                e.Cancel = true;
                return;
            }

            SaveWindowBounds();

            // Clean shutdown: no crash, so discard the recovery draft
            _autosaveTimer?.Stop();
            DeleteDraft();

            // When the last window closes, install a background-downloaded
            // update if one is waiting.
            bool lastWindow = Application.Current.Windows.OfType<MainWindow>().Count() <= 1;
            if (lastWindow)
            {
                var path = _settings.PendingUpdatePath;
                if (!string.IsNullOrEmpty(path) && File.Exists(path))
                {
                    ClearPendingUpdate();
                    UpdateService.LaunchInstaller(path);
                }
            }
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
