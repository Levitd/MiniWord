using System.Linq;
using System.Windows;
using System.Windows.Media;
using MiniWord.Models;
using MiniWord.Services;

namespace MiniWord
{
    public partial class SettingsDialog : Window
    {
        private readonly AppSettings _settings;

        public SettingsDialog(AppSettings settings)
        {
            InitializeComponent();
            _settings = settings;

            FontCombo.ItemsSource = Fonts.SystemFontFamilies
                .Select(f => f.Source)
                .OrderBy(s => s)
                .ToList();
            FontCombo.SelectedItem = _settings.DefaultFontFamily;

            SizeCombo.ItemsSource = new[] { "8", "9", "10", "11", "12", "14", "16", "18", "20", "22", "24", "28", "36", "48", "72" };
            SizeCombo.SelectedItem = _settings.DefaultFontSize.ToString("F0");

            ApplyLocalization();
            LangCombo.SelectedIndex = _settings.Language == "ru" ? 1 : 0;
        }

        private void ApplyLocalization()
        {
            Title = Loc.T("SettingsTitle");
            FontLabel.Content = Loc.T("DefaultFont");
            SizeLabel.Content = Loc.T("DefaultFontSize");
            LangLabel.Content = Loc.T("LanguageLabel");
            NoteText.Text = Loc.T("RestartNote");
            OkButton.Content = Loc.T("OK");
            CancelButton.Content = Loc.T("Cancel");
            LangCombo.ItemsSource = new[] { Loc.T("LangEnglish"), Loc.T("LangRussian") };
        }

        private void OK_Click(object sender, RoutedEventArgs e)
        {
            if (FontCombo.SelectedItem is string font)
                _settings.DefaultFontFamily = font;

            if (SizeCombo.SelectedItem is string sizeStr && double.TryParse(sizeStr, out var size))
                _settings.DefaultFontSize = size;

            _settings.Language = LangCombo.SelectedIndex == 1 ? "ru" : "en";
            _settings.Save();

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
