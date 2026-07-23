using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Windows;
using System.Windows.Navigation;
using MiniWord.Services;

namespace MiniWord
{
    public partial class AboutDialog : Window
    {
        public AboutDialog()
        {
            InitializeComponent();

            Title = Loc.T("AboutTitle");
            DescriptionText.Text = Loc.T("AboutDescription");
            VersionLabel.Text = Loc.T("VersionLabel");
            ReleaseDateLabel.Text = Loc.T("ReleaseDateLabel");
            AuthorLabel.Text = Loc.T("AuthorLabel");
            OkButton.Content = Loc.T("OK");

            var version = Assembly.GetExecutingAssembly().GetName().Version;
            VersionText.Text = version != null ? version.ToString(3) : "1.0";

            // Build timestamp of the running executable serves as the release date
            try
            {
                var exePath = Environment.ProcessPath;
                if (!string.IsNullOrEmpty(exePath) && File.Exists(exePath))
                    ReleaseDateText.Text = File.GetLastWriteTime(exePath).ToString("dd.MM.yyyy");
            }
            catch
            {
                ReleaseDateText.Text = "—";
            }
        }

        private void Link_RequestNavigate(object sender, RequestNavigateEventArgs e)
        {
            Process.Start(new ProcessStartInfo(e.Uri.AbsoluteUri) { UseShellExecute = true });
            e.Handled = true;
        }

        private void OK_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
