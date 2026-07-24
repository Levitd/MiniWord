using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows;

namespace MiniWord
{
    public partial class App : Application
    {
        // Named mutex the installer looks for (AppMutex) so it can update
        // the app in place while a copy is running.
        public const string AppMutexName = "MiniWord_App_Mutex_8F2A1C64";
        private Mutex? _mutex;

        // Extensions MiniWord can open (HTML/PDF are export-only, not listed).
        private static readonly string[] OpenableExtensions = { ".docx", ".rtf", ".odt", ".txt" };

        protected override void OnStartup(StartupEventArgs e)
        {
            _mutex = new Mutex(true, AppMutexName);
            base.OnStartup(e);

            // Open any supported document passed on the command line (file
            // association / "Open with"), each in its own window.
            var files = e.Args
                .Where(a => File.Exists(a) &&
                            OpenableExtensions.Any(ext => a.EndsWith(ext, StringComparison.OrdinalIgnoreCase)))
                .ToList();

            if (files.Count > 0)
            {
                foreach (var file in files)
                    new MainWindow(file).Show();
            }
            else
            {
                // Plain launch: the window offers recent files on its own
                new MainWindow(null).Show();
            }
        }

        protected override void OnExit(ExitEventArgs e)
        {
            _mutex?.Dispose();
            base.OnExit(e);
        }
    }
}
