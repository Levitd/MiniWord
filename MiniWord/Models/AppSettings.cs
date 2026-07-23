using System;
using System.IO;
using System.Text.Json;

namespace MiniWord.Models
{
    public class AppSettings
    {
        public string DefaultFontFamily { get; set; } = "Calibri";
        public double DefaultFontSize { get; set; } = 12;
        public string Language { get; set; } = "en";
        public string PageSize { get; set; } = "A4";

        private static string SettingsPath => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "MiniWord", "settings.json");

        public static AppSettings Load()
        {
            try
            {
                if (File.Exists(SettingsPath))
                {
                    var loaded = JsonSerializer.Deserialize<AppSettings>(File.ReadAllText(SettingsPath));
                    if (loaded != null)
                        return loaded;
                }
            }
            catch { }
            return CreateDefault();
        }

        private static AppSettings CreateDefault()
        {
            var settings = new AppSettings();
            var uiLang = System.Globalization.CultureInfo.CurrentUICulture.TwoLetterISOLanguageName;
            settings.Language = uiLang == "ru" ? "ru" : "en";
            return settings;
        }

        public void Save()
        {
            try
            {
                var dir = Path.GetDirectoryName(SettingsPath);
                if (!string.IsNullOrEmpty(dir))
                    Directory.CreateDirectory(dir);
                File.WriteAllText(SettingsPath, JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true }));
            }
            catch { }
        }
    }
}
