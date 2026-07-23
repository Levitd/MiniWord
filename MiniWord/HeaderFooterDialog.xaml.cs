using System.Windows;
using MiniWord.Models;
using MiniWord.Services;

namespace MiniWord
{
    public partial class HeaderFooterDialog : Window
    {
        private readonly DocumentSettings _settings;

        public HeaderFooterDialog(DocumentSettings settings)
        {
            InitializeComponent();
            _settings = settings;

            Title = Loc.T("HFTitle");
            HeaderLabel.Content = Loc.T("HeaderLabel");
            FooterLabel.Content = Loc.T("FooterLabel");
            PageNumbersCheck.Content = Loc.T("PageNumbersCheck");
            PositionLabel.Content = Loc.T("PositionLabel");
            OkButton.Content = Loc.T("OK");
            CancelButton.Content = Loc.T("Cancel");
            PositionCombo.ItemsSource = new[]
            {
                Loc.T("PosFooterCenter"),
                Loc.T("PosFooterRight"),
                Loc.T("PosHeaderRight")
            };

            HeaderBox.Text = _settings.HeaderText;
            FooterBox.Text = _settings.FooterText;
            PageNumbersCheck.IsChecked = _settings.ShowPageNumbers;
            PositionCombo.SelectedIndex = (int)_settings.PageNumberPosition;
        }

        private void OK_Click(object sender, RoutedEventArgs e)
        {
            _settings.HeaderText = HeaderBox.Text.Trim();
            _settings.FooterText = FooterBox.Text.Trim();
            _settings.ShowPageNumbers = PageNumbersCheck.IsChecked == true;
            _settings.PageNumberPosition = (PageNumberPosition)System.Math.Max(0, PositionCombo.SelectedIndex);

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
