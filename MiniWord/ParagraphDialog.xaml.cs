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
            WireValueLabels();
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

        private void WireValueLabels()
        {
            Bind(LeftIndentSlider, LeftIndentValue);
            Bind(RightIndentSlider, RightIndentValue);
            Bind(FirstLineIndentSlider, FirstLineIndentValue);
            Bind(LineSpacingSlider, LineSpacingValue);
            Bind(SpaceBeforeSlider, SpaceBeforeValue);
            Bind(SpaceAfterSlider, SpaceAfterValue);
        }

        private static void Bind(Slider slider, TextBlock label)
        {
            slider.ValueChanged += (s, e) => label.Text = slider.Value.ToString("F0");
            label.Text = slider.Value.ToString("F0");
        }

        private void LoadCurrentSettings()
        {
            var para = _textEditor.Selection.Start?.Paragraph;
            if (para == null)
                return;

            LeftIndentSlider.Value = DipToPt(para.Margin.Left);
            RightIndentSlider.Value = DipToPt(para.Margin.Right);
            SpaceBeforeSlider.Value = DipToPt(para.Margin.Top);
            SpaceAfterSlider.Value = DipToPt(para.Margin.Bottom);
            FirstLineIndentSlider.Value = DipToPt(para.TextIndent);

            if (para.LineHeight > 0 && !double.IsNaN(para.LineHeight))
                LineSpacingSlider.Value = DipToPt(para.LineHeight);
        }

        private void OK_Click(object sender, RoutedEventArgs e)
        {
            var para = _textEditor.Selection.Start?.Paragraph;
            if (para != null)
            {
                para.Margin = new Thickness(
                    PtToDip(LeftIndentSlider.Value), PtToDip(SpaceBeforeSlider.Value),
                    PtToDip(RightIndentSlider.Value), PtToDip(SpaceAfterSlider.Value));
                para.TextIndent = PtToDip(FirstLineIndentSlider.Value);
                if (LineSpacingSlider.Value > 0)
                    para.LineHeight = PtToDip(LineSpacingSlider.Value);
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
