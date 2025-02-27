using System;
using System.IO;
using ExcelProcessor.Interfaces;

namespace ExcelProcessor.Services
{
    public class LogService : ILogService
    {
        private readonly string _logFilePath;

        public LogService(string logFilePath = "app.log")
        {
            _logFilePath = logFilePath;
        }

        public void LogInformation(string message)
        {
            Log($"INFO: {message}");
        }

        public void LogError(string message, Exception ex = null)
        {
            Log($"ERROR: {message}");
            if (ex != null)
            {
                Log($"Exception: {ex}");
            }
        }

        private void Log(string message)
        {
            try
            {
                File.AppendAllText(_logFilePath, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}{Environment.NewLine}");
            }
            catch
            {
                // Suppress logging errors
            }
        }
    }
}
