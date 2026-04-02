using System;
using System.Diagnostics;

namespace Connections.Services.Core
{
    /// <summary>
    /// Simple logger that writes to System.Diagnostics.Debug.
    /// </summary>
    public class DebugLogger : ILogger
    {
        public void Debug(string message)
        {
            System.Diagnostics.Debug.WriteLine($"[DEBUG] {message}");
        }

        public void Info(string message)
        {
            System.Diagnostics.Debug.WriteLine($"[INFO] {message}");
        }

        public void Warning(string message)
        {
            System.Diagnostics.Debug.WriteLine($"[WARN] {message}");
        }

        public void Error(string message, Exception ex = null)
        {
            System.Diagnostics.Debug.WriteLine($"[ERROR] {message}");
            if (ex != null)
            {
                System.Diagnostics.Debug.WriteLine(ex.ToString());
            }
        }
    }
}
