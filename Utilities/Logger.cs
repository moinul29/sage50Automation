namespace Sage50Automation.Utilities
{
    /// <summary>
    /// Centralized logging utility for test execution.
    /// 
    /// Writes to three destinations simultaneously:
    ///   1. Console       — visible in terminal with -v d flag
    ///   2. Memory buffer — available within the same process
    ///   3. Temp file     — persists across separate test processes
    /// 
    /// Usage:
    ///   var logger = new Logger(logFilePath);
    ///   logger.Info("Step 1: Launching application...");
    ///   logger.Clear();  // Call at start of first test only
    /// </summary>
    public class Logger
    {
        private static readonly System.Text.StringBuilder _buffer = new();
        private readonly string _logFilePath;

        public Logger(string logFilePath)
        {
            _logFilePath = logFilePath;
        }

        /// <summary>
        /// Write a message to console, memory buffer, and persistent file
        /// </summary>
        public void Info(string message)
        {
            Console.WriteLine(message);
            _buffer.AppendLine(message);
            try
            {
                File.AppendAllText(_logFilePath, message + Environment.NewLine);
            }
            catch
            {
                // Ignore file write errors (file may be locked by another process)
            }
        }

        /// <summary>
        /// Clear all logs — call at the start of the FIRST test run (Actian 2026)
        /// </summary>
        public void Clear()
        {
            _buffer.Clear();
            try
            {
                if (File.Exists(_logFilePath))
                    File.Delete(_logFilePath);
            }
            catch
            {
                // Ignore deletion errors
            }
        }

        /// <summary>
        /// Get all accumulated logs (reads from file first, falls back to memory)
        /// </summary>
        public string GetAllLogs()
        {
            try
            {
                if (File.Exists(_logFilePath))
                    return File.ReadAllText(_logFilePath);
            }
            catch { }
            return _buffer.ToString();
        }

        /// <summary>
        /// Path to the persistent log file
        /// </summary>
        public string LogFilePath => _logFilePath;
    }
}
