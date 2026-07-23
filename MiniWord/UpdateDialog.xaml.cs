using System;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using MiniWord.Models;
using MiniWord.Services;

namespace MiniWord
{
    public partial class UpdateDialog : Window
    {
        private readonly UpdateInfo _info;
        private bool _busy;

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
            BackgroundButton.Content = Loc.T("UpdateInBackground");
            LaterButton.Content = Loc.T("UpdateLater");
        }

        private async void Update_Click(object sender, RoutedEventArgs e)
        {
            if (_busy)
                return;
            _busy = true;

            UpdateButton.IsEnabled = false;
            BackgroundButton.IsEnabled = false;
            LaterButton.IsEnabled = false;
            ProgressPanel.Visibility = Visibility.Visible;

            try
            {
                var progress = new Progress<double>(p => DownloadProgress.Value = p);
                var installerPath = await UpdateService.DownloadInstallerAsync(_info, progress);

                UpdateService.LaunchInstaller(installerPath);
                Application.Current.Shutdown();
            }
            catch
            {
                MessageBox.Show(Loc.T("UpdateDownloadError"), Loc.T("UpdateTitle"),
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                _busy = false;
                UpdateButton.IsEnabled = true;
                BackgroundButton.IsEnabled = true;
                LaterButton.IsEnabled = true;
                ProgressPanel.Visibility = Visibility.Collapsed;
            }
        }

        // Download silently in the background; the app keeps working. When
        // done, the update is remembered and installed on close / next launch.
        private void Background_Click(object sender, RoutedEventArgs e)
        {
            var info = _info;
            _ = Task.Run(async () =>
            {
                try
                {
                    var path = await UpdateService.DownloadInstallerAsync(info, null);
                    var settings = AppSettings.Current;
                    settings.PendingUpdatePath = path;
                    settings.PendingUpdateVersion = info.Version.ToString(3);
                    settings.Save();
                }
                catch
                {
                    // Silent: a failed background download just means no pending update
                }
            });

            Close();
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
