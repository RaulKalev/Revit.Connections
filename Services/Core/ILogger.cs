using System;

namespace Connections.Services.Core
{
    /// <summary>
    /// Interface for logging throughout the application.
    /// </summary>
    public interface ILogger
    {
        void Debug(string message);
        void Info(string message);
        void Warning(string message);
        void Error(string message, Exception ex = null);
    }
}
