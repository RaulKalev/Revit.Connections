using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using System;
using System.Runtime.InteropServices;

namespace Connections.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class ConnectionsCommand : IExternalCommand
    {
        private static UI.ConnectionsWindow _window;
        private static bool _pendingShow;

        [DllImport("user32.dll")] private static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll")] private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        private const int SW_RESTORE = 9;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                if (_window != null && _window.IsLoaded)
                {
                    var hwnd = new System.Windows.Interop.WindowInteropHelper(_window).Handle;
                    if (_window.WindowState == System.Windows.WindowState.Minimized)
                        ShowWindow(hwnd, SW_RESTORE);

                    _window.Activate();
                    _window.Focus();
                    SetForegroundWindow(hwnd);
                    return Result.Succeeded;
                }

                try { var dummy = new MaterialDesignThemes.Wpf.PaletteHelper(); } catch { }

                if (!_pendingShow)
                {
                    _pendingShow = true;
                    commandData.Application.Idling += OnRevitIdling;
                }

                return Result.Succeeded;
            }
            catch (System.Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        private static void OnRevitIdling(object sender, IdlingEventArgs e)
        {
            var uiApp = sender as UIApplication;
            if (uiApp != null)
                uiApp.Idling -= OnRevitIdling;

            _pendingShow = false;

            try
            {
                if (_window != null && _window.IsLoaded)
                {
                    _window.Activate();
                    return;
                }

                _window = new UI.ConnectionsWindow(uiApp, App.ExternalEventService);
                var owner = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle;
                new System.Windows.Interop.WindowInteropHelper(_window) { Owner = owner };

                _window.Closed += (s, ev) => { _window = null; };
                _window.Show();
            }
            catch (Exception ex)
            {
                App.Logger?.Error("Failed to create Connections window in Idling handler", ex);
            }
        }
    }
}
