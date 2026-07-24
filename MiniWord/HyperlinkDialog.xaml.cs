using System.Windows;
using MiniWord.Services;

namespace MiniWord
{
    public partial class HyperlinkDialog : Window
    {
        public string DisplayText => string.IsNullOrWhiteSpace(TextBox.Text) ? UrlBox.Text : TextBox.Text;
        public string Url => UrlBox.Text.Trim();

        public HyperlinkDialog(string selectedText)
        {
            InitializeComponent();

            Title = Loc.T("HyperlinkTitle");
            TextLabel.Content = Loc.T("LinkTextLabel");
            UrlLabel.Content = Loc.T("LinkUrlLabel");
            OkButton.Content = Loc.T("OK");
            CancelButton.Content = Loc.T("Cancel");

            TextBox.Text = selectedText ?? "";
            UrlBox.Focus();
        }

        private void OK_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(UrlBox.Text))
                return;
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
