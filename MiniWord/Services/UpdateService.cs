using System;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Reflection;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace MiniWord.Services
{
    public class UpdateInfo
    {
        public Version Version { get; init; } = new(0, 0);
        public string TagName { get; init; } = "";
        public string Notes { get; init; } = "";
        public string InstallerUrl { get; init; } = "";
        public string InstallerName { get; init; } = "";
    }

    /// <summary>
    /// Checks the GitHub Releases API for a newer version and downloads the
    /// installer. No third-party dependencies — plain HttpClient + JSON.
    /// </summary>
    public static class UpdateService
    {
        private const string LatestReleaseApi =
            "https://api.github.com/repos/Levitd/MiniWord/releases/latest";

        private static readonly HttpClient Http = CreateClient();

        private static HttpClient CreateClient()
        {
            var client = new HttpClient();
            // GitHub API rejects requests without a User-Agent
            client.DefaultRequestHeaders.UserAgent.ParseAdd("MiniWord-Updater");
            client.DefaultRequestHeaders.Accept.ParseAdd("application/vnd.github+json");
            // Kept short so a dead/slow network never delays things noticeably;
            // the check is non-blocking anyway (the window is already shown).
            client.Timeout = TimeSpan.FromSeconds(15);
            return client;
        }

        public static Version CurrentVersion =>
            Assembly.GetExecutingAssembly().GetName().Version ?? new Version(1, 0, 0);

        /// <summary>
        /// Returns update info if the latest release is newer than the running
        /// version, otherwise null. Never throws — network errors return null.
        /// </summary>
        public static Task<UpdateInfo?> CheckForUpdateAsync(CancellationToken ct = default) =>
            CheckForUpdateAsync(LatestReleaseApi, ct);

        public static async Task<UpdateInfo?> CheckForUpdateAsync(string apiUrl, CancellationToken ct = default)
        {
            try
            {
                var json = await Http.GetStringAsync(apiUrl, ct);
                using var doc = JsonDocument.Parse(json);
                var root = doc.RootElement;

                var tag = root.GetProperty("tag_name").GetString() ?? "";
                var version = ParseVersion(tag);
                if (version == null || version <= CurrentVersion)
                    return null;

                string notes = root.TryGetProperty("body", out var b) ? b.GetString() ?? "" : "";

                string installerUrl = "", installerName = "";
                if (root.TryGetProperty("assets", out var assets))
                {
                    foreach (var asset in assets.EnumerateArray())
                    {
                        var name = asset.GetProperty("name").GetString() ?? "";
                        if (name.EndsWith(".exe", StringComparison.OrdinalIgnoreCase))
                        {
                            installerName = name;
                            installerUrl = asset.GetProperty("browser_download_url").GetString() ?? "";
                            break;
                        }
                    }
                }

                if (string.IsNullOrEmpty(installerUrl))
                    return null;

                return new UpdateInfo
                {
                    Version = version,
                    TagName = tag,
                    Notes = notes,
                    InstallerUrl = installerUrl,
                    InstallerName = installerName
                };
            }
            catch
            {
                return null;
            }
        }

        private static string UpdatesDir
        {
            get
            {
                var dir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "MiniWord", "updates");
                Directory.CreateDirectory(dir);
                return dir;
            }
        }

        /// <summary>
        /// Downloads the installer to a persistent folder, reporting 0..1
        /// progress. Returns the local path.
        /// </summary>
        public static async Task<string> DownloadInstallerAsync(
            UpdateInfo info, IProgress<double>? progress, CancellationToken ct = default)
        {
            var targetPath = Path.Combine(UpdatesDir, info.InstallerName);

            using var response = await Http.GetAsync(
                info.InstallerUrl, HttpCompletionOption.ResponseHeadersRead, ct);
            response.EnsureSuccessStatusCode();

            var total = response.Content.Headers.ContentLength ?? -1L;
            await using var source = await response.Content.ReadAsStreamAsync(ct);
            await using var dest = new FileStream(targetPath, FileMode.Create, FileAccess.Write, FileShare.None);

            var buffer = new byte[81920];
            long read = 0;
            int n;
            while ((n = await source.ReadAsync(buffer, ct)) > 0)
            {
                await dest.WriteAsync(buffer.AsMemory(0, n), ct);
                read += n;
                if (total > 0)
                    progress?.Report((double)read / total);
            }

            return targetPath;
        }

        /// <summary>
        /// Launches the installer after a short delay so this process has time
        /// to exit and release MiniWord.exe, then quits nothing itself — the
        /// caller is responsible for shutting the app down.
        /// </summary>
        public static void LaunchInstaller(string installerPath)
        {
            Process.Start(new ProcessStartInfo("cmd.exe")
            {
                Arguments = $"/c timeout /t 3 /nobreak >nul & \"{installerPath}\" /SILENT",
                UseShellExecute = false,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden
            });
        }

        private static Version? ParseVersion(string tag)
        {
            var t = tag.TrimStart('v', 'V').Trim();
            return Version.TryParse(t, out var v) ? v : null;
        }
    }
}
