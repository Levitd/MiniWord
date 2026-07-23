using System;
using System.Diagnostics;
using System.Text.RegularExpressions;
using System.Windows;
using MiniWord.Services;

namespace MiniWord
{
    public partial class UpdateDialog : Window
    {
        private readonly UpdateInfo _info;
        private bool _downloading;

        public UpdateDialog(UpdateInfo info)
        {
            InitializeComponent();
            _info = info;

            Title = Loc.T("UpdateTitle");
            HeadlineText.Text = string.Format(Loc.T("UpdateText"),
                info.Version.ToString(3), UpdateService.CurrentVersion.ToString(3));
            WhatsNewLabel.Text = Loc.T("WhatsNew");
            NotesText.Text = string.IsNullOrWhiteSpace(info.Notes) ? "—" : StripMarkdown(info.Notes);
            ProgressLabel.Text = Loc.T("Downloading");
            UpdateButton.Content = Loc.T("UpdateNow");
            LaterButton.Content = Loc.T("UpdateLater");
        }

        private async void Update_Click(object sender, RoutedEventArgs e)
        {
            if (_downloading)
                return;
            _downloading = true;

            UpdateButton.IsEnabled = false;
            LaterButton.IsEnabled = false;
            ProgressPanel.Visibility = Visibility.Visible;

            try
            {
                var progress = new Progress<double>(p => DownloadProgress.Value = p);
                var installerPath = await UpdateService.DownloadInstallerAsync(_info, progress);

                // The installer cannot overwrite MiniWord.exe while this
                // process still holds it, so start it through a short delay
                // and quit immediately — by the time it runs, we are gone.
                Process.Start(new ProcessStartInfo("cmd.exe")
                {
                    Arguments = $"/c timeout /t 3 /nobreak >nul & \"{installerPath}\" /SILENT",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    WindowStyle = ProcessWindowStyle.Hidden
                });

                Application.Current.Shutdown();
            }
            catch
            {
                MessageBox.Show(Loc.T("UpdateDownloadError"), Loc.T("UpdateTitle"),
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                _downloading = false;
                UpdateButton.IsEnabled = true;
                LaterButton.IsEnabled = true;
                ProgressPanel.Visibility = Visibility.Collapsed;
            }
        }

        /// GitHub release notes are Markdown; show them as plain text.
        private static string StripMarkdown(string text)
        {
            text = Regex.Replace(text, @"\*\*(.+?)\*\*", "$1");   // bold
            text = Regex.Replace(text, @"(?m)^\s*#{1,6}\s*", "");  // headings
            text = Regex.Replace(text, @"(?m)^\s*---+\s*$", "");   // rules
            text = Regex.Replace(text, @"\[(.+?)\]\((.+?)\)", "$1 ($2)"); // links
            return text.Trim();
        }

        private void Later_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
