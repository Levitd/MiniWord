using System.Windows;
using System.Windows.Controls;
using MiniWord.Services;

namespace MiniWord
{
    public partial class ParagraphDialog : Window
    {
        private readonly RichTextBox _textEditor;

        // Dialog values are in points; WPF paragraph properties are in DIP (1/96")
        private static double PtToDip(double pt) => pt * 96.0 / 72.0;
        private static double DipToPt(double dip) => dip * 72.0 / 96.0;

        public ParagraphDialog(RichTextBox textEditor)
        {
            InitializeComponent();
            _textEditor = textEditor;
            ApplyLocalization();
            LoadCurrentSettings();
        }

        private void ApplyLocalization()
        {
            Title = Loc.T("Paragraph");
            MainGroup.Header = Loc.T("IndentSpacing");
            LeftIndentLabel.Content = Loc.T("LeftIndent");
            RightIndentLabel.Content = Loc.T("RightIndent");
            FirstLineIndentLabel.Content = Loc.T("FirstLineIndent");
            LineSpacingLabel.Content = Loc.T("LineSpacing");
            SpaceBeforeLabel.Content = Loc.T("SpaceBefore");
            SpaceAfterLabel.Content = Loc.T("SpaceAfter");
            OkButton.Content = Loc.T("OK");
            CancelButton.Content = Loc.T("Cancel");
        }

        private void LoadCurrentSettings()
        {
            var para = _textEditor.Selection.Start?.Paragraph;
            if (para == null)
                return;

            LeftIndentBox.Text = DipToPt(para.Margin.Left).ToString("F0");
            RightIndentBox.Text = DipToPt(para.Margin.Right).ToString("F0");
            SpaceBeforeBox.Text = DipToPt(para.Margin.Top).ToString("F0");
            SpaceAfterBox.Text = DipToPt(para.Margin.Bottom).ToString("F0");
            FirstLineIndentBox.Text = DipToPt(para.TextIndent).ToString("F0");

            if (para.LineHeight > 0 && !double.IsNaN(para.LineHeight))
                LineSpacingBox.Text = DipToPt(para.LineHeight).ToString("F0");
        }

        private void OK_Click(object sender, RoutedEventArgs e)
        {
            double.TryParse(LeftIndentBox.Text, out var leftIndent);
            double.TryParse(RightIndentBox.Text, out var rightIndent);
            double.TryParse(FirstLineIndentBox.Text, out var firstLineIndent);
            double.TryParse(SpaceBeforeBox.Text, out var spaceBefore);
            double.TryParse(SpaceAfterBox.Text, out var spaceAfter);
            double.TryParse(LineSpacingBox.Text, out var lineSpacing);

            var para = _textEditor.Selection.Start?.Paragraph;
            if (para != null)
            {
                para.Margin = new Thickness(
                    PtToDip(leftIndent), PtToDip(spaceBefore),
                    PtToDip(rightIndent), PtToDip(spaceAfter));
                para.TextIndent = PtToDip(firstLineIndent);
                if (lineSpacing > 0)
                    para.LineHeight = PtToDip(lineSpacing);
            }

            DialogResult = true;
            Close();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
