using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using MiniWord.Services;

namespace MiniWord
{
    public partial class StartDialog : Window
    {
        public string? SelectedFile { get; private set; }

        private class RecentEntry
        {
            public string FullPath { get; init; } = "";
            public string FileName => Path.GetFileName(FullPath);
        }

        public StartDialog(IEnumerable<string> recentFiles)
        {
            InitializeComponent();

            Title = "MiniWord";
            RecentLabel.Text = Loc.T("RecentFilesLabel");
            OpenSelectedButton.Content = Loc.T("Open").Replace("_", "").Replace(".", "");
            BrowseButton.Content = Loc.T("OpenOther");
            NewButton.Content = Loc.T("NewDocument");

            RecentList.ItemsSource = recentFiles
                .Where(File.Exists)
                .Select(p => new RecentEntry { FullPath = p })
                .ToList();
            if (RecentList.Items.Count > 0)
                RecentList.SelectedIndex = 0;
        }

        private void OpenSelected_Click(object sender, RoutedEventArgs e)
        {
            if (RecentList.SelectedItem is RecentEntry entry)
            {
                SelectedFile = entry.FullPath;
                DialogResult = true;
                Close();
            }
        }

        private void RecentList_DoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (RecentList.SelectedItem is RecentEntry)
                OpenSelected_Click(sender, e);
        }

        private void Browse_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog { Filter = Loc.T("FilterDocx"), DefaultExt = ".docx" };
            if (dlg.ShowDialog() == true)
            {
                SelectedFile = dlg.FileName;
                DialogResult = true;
                Close();
            }
        }

        private void New_Click(object sender, RoutedEventArgs e)
        {
            SelectedFile = null;
            DialogResult = false;
            Close();
        }
    }
}
