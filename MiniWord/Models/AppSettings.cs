using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;

namespace MiniWord.Models
{
    public class AppSettings
    {
        public const int MaxRecentFiles = 5;

        public string DefaultFontFamily { get; set; } = "Calibri";
        public double DefaultFontSize { get; set; } = 12;
        public string Language { get; set; } = "en";
        public string PageSize { get; set; } = "A4";

        // Window bounds from the last session (0 / null = not saved yet)
        public double WindowWidth { get; set; }
        public double WindowHeight { get; set; }
        public double? WindowLeft { get; set; }
        public double? WindowTop { get; set; }
        public bool WindowMaximized { get; set; }

        // Most-recently-opened files, newest first
        public List<string> RecentFiles { get; set; } = new();

        // A background-downloaded update waiting to be installed
        public string? PendingUpdatePath { get; set; }
        public string? PendingUpdateVersion { get; set; }

        private static string SettingsPath => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "MiniWord", "settings.json");

        // Process-wide shared instance so every window sees the same recent
        // files, window bounds and pending update.
        private static AppSettings? _current;
        public static AppSettings Current => _current ??= Load();

        public static AppSettings Load()
        {
            try
            {
                if (File.Exists(SettingsPath))
                {
                    var loaded = JsonSerializer.Deserialize<AppSettings>(File.ReadAllText(SettingsPath));
                    if (loaded != null)
                    {
                        loaded.RecentFiles ??= new List<string>();
                        return loaded;
                    }
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

        public void AddRecentFile(string path)
        {
            if (string.IsNullOrEmpty(path))
                return;

            RecentFiles.RemoveAll(p => string.Equals(p, path, StringComparison.OrdinalIgnoreCase));
            RecentFiles.Insert(0, path);
            if (RecentFiles.Count > MaxRecentFiles)
                RecentFiles = RecentFiles.Take(MaxRecentFiles).ToList();
            Save();
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
