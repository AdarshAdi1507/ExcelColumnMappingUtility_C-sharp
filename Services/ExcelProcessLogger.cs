// ExcelProcessLogger.cs
using System;
using System.IO;
using System.Text;
using System.Diagnostics;
using ExcelProcessor.Interfaces;

namespace ExcelProcessor.Services
{
    public class ExcelProcessLogger : ILogService
    {
        private string _logFolderPath;
        private string _currentLogPath;
        private readonly StringBuilder _logBuilder;
        private readonly Stopwatch _timer;
        private int _processedRows;
        private bool _isInitialized;

        public ExcelProcessLogger()
        {
            _logBuilder = new StringBuilder();
            _timer = new Stopwatch();
            _processedRows = 0;
            _isInitialized = false;
        }

        private void InitializeLogPaths(string excelFilePath)
        {
            if (!_isInitialized)
            {
                // Create logs directory in the same folder as the Excel file
                string excelDirectory = Path.GetDirectoryName(excelFilePath);
                _logFolderPath = Path.Combine(excelDirectory, "ConversionLogs");

                // Create the directory if it doesn't exist
                if (!Directory.Exists(_logFolderPath))
                {
                    Directory.CreateDirectory(_logFolderPath);
                }

                // Create log file with timestamp
                string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                string excelFileName = Path.GetFileNameWithoutExtension(excelFilePath);
                _currentLogPath = Path.Combine(_logFolderPath, $"{excelFileName}_Conversion_{timestamp}.txt");

                _isInitialized = true;

                // Write initial log entry to verify file creation
                LogInformation($"Log file initialized at: {_currentLogPath}");
            }
        }

        public void StartProcess(string inputFilePath, string selectedSheet)
        {
            InitializeLogPaths(inputFilePath);
            _timer.Restart();
            _processedRows = 0;

            LogInformation("=== Excel Conversion Process Started ===");
            LogInformation($"Start Time: {DateTime.Now}");
            LogInformation($"Input File: {inputFilePath}");
            LogInformation($"Selected Sheet: {selectedSheet}");
            LogInformation($"Log File Location: {_currentLogPath}");
        }

        public void LogRowProcessed()
        {
            _processedRows++;
            if (_processedRows % 1000 == 0) // Log every 100 rows to avoid too frequent writes
            {
                LogInformation($"Processed {_processedRows} rows...");
            }
        }

        public void EndProcess(string outputPath)
        {
            _timer.Stop();
            LogInformation("\n=== Process Summary ===");
            LogInformation($"Process Completed at: {DateTime.Now}");
            LogInformation($"Total Time Taken: {_timer.Elapsed.ToString(@"hh\:mm\:ss\.fff")}");
            LogInformation($"Total Rows Processed: {_processedRows}");
            LogInformation($"Output File Location: {outputPath}");
            LogInformation("=== Process Completed Successfully ===\n");
        }

        public void LogInformation(string message)
        {
            string logMessage = $"[INFO] {DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} - {message}";
            _logBuilder.AppendLine(logMessage);
            WriteToFile(logMessage);
        }

        public void LogError(string message, Exception ex = null)
        {
            string logMessage = $"[ERROR] {DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} - {message}";
            if (ex != null)
            {
                logMessage += $"\nException: {ex.Message}";
                logMessage += $"\nStack Trace: {ex.StackTrace}";
            }
            _logBuilder.AppendLine(logMessage);
            WriteToFile(logMessage);
        }

        private void WriteToFile(string message)
        {
            if (string.IsNullOrEmpty(_currentLogPath))
            {
                return; // Skip if log path not initialized
            }

            try
            {
                // Ensure directory exists before writing
                Directory.CreateDirectory(Path.GetDirectoryName(_currentLogPath));

                // Use FileStream to ensure proper file handling
                using (FileStream fs = new FileStream(_currentLogPath, FileMode.Append, FileAccess.Write, FileShare.Read))
                using (StreamWriter sw = new StreamWriter(fs))
                {
                    sw.WriteLine(message);
                }
            }
            catch (Exception ex)
            {
                // In case of error, try to write to the application directory as fallback
                string fallbackPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "error_log.txt");
                try
                {
                    File.AppendAllText(fallbackPath, $"Error writing to log: {ex.Message}\n{message}\n");
                }
                catch
                {
                    // Suppress fallback logging errors
                }
            }
        }

        public string GetLogContent()
        {
            return _logBuilder.ToString();
        }

        public string GetLogFilePath()
        {
            return _currentLogPath;
        }
    }
}