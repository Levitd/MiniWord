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
        }

        protected override void OnExit(ExitEventArgs e)
        {
            _mutex?.Dispose();
            base.OnExit(e);
        }
    }
}
