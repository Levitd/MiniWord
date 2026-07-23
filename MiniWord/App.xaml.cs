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

        protected override void OnStartup(StartupEventArgs e)
        {
            _mutex = new Mutex(true, AppMutexName);
            base.OnStartup(e);

            // Open any .docx paths passed on the command line (file association
            // / "Open with"), each in its own window.
            var files = e.Args
                .Where(a => File.Exists(a) &&
                            a.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
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
