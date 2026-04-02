using Autodesk.Revit.UI;
using ricaun.Revit.UI;
using System;
using System.IO;
using Connections.Commands;

namespace Connections
{
    [AppLoader]
    public class App : IExternalApplication
    {
        public static Services.Revit.RevitExternalEventService ExternalEventService { get; private set; }
        public static Services.Core.SessionLogger Logger { get; private set; }
        private RibbonPanel ribbonPanel;

        private static readonly string CrashLogPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "RK Tools", "Connections", "crash.log");

        public Result OnStartup(UIControlledApplication application)
        {
            Logger = new Services.Core.SessionLogger();
            ExternalEventService = new Services.Revit.RevitExternalEventService(Logger);

            // Crash diagnostics
            AppDomain.CurrentDomain.UnhandledException += (s, e) =>
            {
                LogCrash("AppDomain.UnhandledException", e.ExceptionObject as Exception);
            };

            if (System.Windows.Application.Current != null)
            {
                System.Windows.Application.Current.DispatcherUnhandledException += (s, e) =>
                {
                    LogCrash("Dispatcher.UnhandledException", e.Exception);
                    e.Handled = true;
                };
            }

            string tabName = "RK Tools";
            try { application.CreateRibbonTab(tabName); } catch { }

            ribbonPanel = application.CreateOrSelectPanel(tabName, "Tools");

            string iconName;
            try
            {
                iconName = UIThemeManager.CurrentTheme == UITheme.Dark
                    ? "Light%20-%20Renumber.tiff"
                    : "Dark%20-%20Renumber.tiff";
            }
            catch
            {
                iconName = "Renumber.tiff";
            }

            ribbonPanel.CreatePushButton<ConnectionsCommand>()
                .SetLargeImage($"pack://application:,,,/Connections;component/Assets/{iconName}")
                .SetText("Connections")
                .SetToolTip("Electrical Connections")
                .SetLongDescription("Connect electrical elements to panels and manage circuit parameters.");

            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            ribbonPanel?.Remove();
            return Result.Succeeded;
        }

        public static void LogCrash(string source, Exception ex)
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(CrashLogPath));
                string entry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] SOURCE: {source}\n" +
                               $"TYPE: {ex?.GetType().FullName}\n" +
                               $"MESSAGE: {ex?.Message}\n" +
                               $"STACK:\n{ex?.StackTrace}\n" +
                               $"INNER: {ex?.InnerException?.Message}\n" +
                               $"INNER STACK:\n{ex?.InnerException?.StackTrace}\n" +
                               new string('=', 80) + "\n";
                File.AppendAllText(CrashLogPath, entry);
            }
            catch { }
        }
    }
}
