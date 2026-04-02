using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;

namespace Connections.Services.Core
{
    /// <summary>
    /// Logger implementation that writes to Debug output AND retains entries
    /// in an in-memory ring buffer. The ring buffer is accessible for optional
    /// UI display or diagnostics.
    ///
    /// Thread safety: Uses a lock to protect the ring buffer. All log methods
    /// are safe to call from any thread.
    ///
    /// Capacity: Retains the most recent MaxEntries entries (default 500).
    /// Older entries are evicted when the buffer is full.
    /// </summary>
    public class SessionLogger : ILogger
    {
        /// <summary>Maximum number of log entries to retain.</summary>
        private const int MaxEntries = 500;

        private readonly List<LogEntry> _entries = new List<LogEntry>();
        private readonly object _lock = new object();

        // ---- ILogger Implementation ----

        public void Debug(string message)
        {
            AddEntry(LogLevel.Debug, message);
            System.Diagnostics.Debug.WriteLine($"[DALI:DEBUG] {message}");
        }

        public void Info(string message)
        {
            AddEntry(LogLevel.Info, message);
            System.Diagnostics.Debug.WriteLine($"[DALI:INFO] {message}");
        }

        public void Warning(string message)
        {
            AddEntry(LogLevel.Warning, message);
            System.Diagnostics.Debug.WriteLine($"[DALI:WARN] {message}");
        }

        public void Error(string message, Exception ex = null)
        {
            string detail = ex != null ? $"{message} | {ex}" : message;
            AddEntry(LogLevel.Error, detail);
            System.Diagnostics.Debug.WriteLine($"[DALI:ERROR] {detail}");
        }

        // ---- Ring Buffer Access ----

        /// <summary>
        /// Returns a snapshot of all retained log entries, oldest first.
        /// Safe to call from any thread.
        /// </summary>
        public List<LogEntry> GetEntries()
        {
            lock (_lock)
            {
                return new List<LogEntry>(_entries);
            }
        }

        /// <summary>
        /// Returns the most recent N entries.
        /// </summary>
        public List<LogEntry> GetRecentEntries(int count)
        {
            lock (_lock)
            {
                int start = Math.Max(0, _entries.Count - count);
                return _entries.GetRange(start, _entries.Count - start);
            }
        }

        /// <summary>
        /// Clears all retained log entries.
        /// </summary>
        public void Clear()
        {
            lock (_lock)
            {
                _entries.Clear();
            }
        }

        // ---- Internal ----

        private void AddEntry(LogLevel level, string message)
        {
            var entry = new LogEntry
            {
                Timestamp = DateTime.Now,
                Level = level,
                Message = message
            };

            lock (_lock)
            {
                if (_entries.Count >= MaxEntries)
                {
                    _entries.RemoveAt(0); // evict oldest
                }
                _entries.Add(entry);
            }
        }
    }

    /// <summary>Log severity level.</summary>
    public enum LogLevel
    {
        Debug,
        Info,
        Warning,
        Error
    }

    /// <summary>A single log entry with timestamp, level, and message.</summary>
    public class LogEntry
    {
        public DateTime Timestamp { get; set; }
        public LogLevel Level { get; set; }
        public string Message { get; set; }

        public override string ToString()
        {
            return $"[{Timestamp:HH:mm:ss}] [{Level}] {Message}";
        }
    }
}
