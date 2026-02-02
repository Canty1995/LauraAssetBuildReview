using System.Collections.Concurrent;
using System.IO;
using System.Text;

namespace LauraAssetBuildReview.Services;

public enum LogLevel
{
    Info,
    Warning,
    Error
}

public class LoggingService
{
    private readonly ConcurrentQueue<string> _logBuffer = new();
    private readonly object _lockObject = new();

    /// <summary>
    /// Logs a message with the specified level.
    /// Messages are buffered for both UI display and file output.
    /// </summary>
    public void Log(string message, LogLevel level = LogLevel.Info)
    {
        var timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        var levelPrefix = level switch
        {
            LogLevel.Warning => "[WARNING]",
            LogLevel.Error => "[ERROR]",
            _ => "[INFO]"
        };

        var logEntry = $"{timestamp} {levelPrefix} {message}";
        
        lock (_lockObject)
        {
            _logBuffer.Enqueue(logEntry);
        }
    }

    /// <summary>
    /// Gets all buffered log messages and clears the buffer.
    /// </summary>
    public List<string> GetAndClearLogs()
    {
        var logs = new List<string>();
        
        lock (_lockObject)
        {
            while (_logBuffer.TryDequeue(out var log))
            {
                logs.Add(log);
            }
        }

        return logs;
    }

    /// <summary>
    /// Gets all buffered log messages without clearing the buffer.
    /// </summary>
    public List<string> GetAllLogs()
    {
        var logs = new List<string>();
        
        lock (_lockObject)
        {
            foreach (var log in _logBuffer)
            {
                logs.Add(log);
            }
        }

        return logs;
    }

    /// <summary>
    /// Writes all buffered logs to a timestamped file next to the main Excel file.
    /// </summary>
    public void FlushToFile(string mainFilePath)
    {
        var directory = Path.GetDirectoryName(mainFilePath) ?? string.Empty;
        var fileName = Path.GetFileNameWithoutExtension(mainFilePath);
        var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
        var logFilePath = Path.Combine(directory, $"{fileName}_EANMatchLog_{timestamp}.txt");

        var allLogs = GetAllLogs();
        if (allLogs.Count == 0)
            return;

        try
        {
            var sb = new StringBuilder();
            sb.AppendLine("=".PadRight(80, '='));
            sb.AppendLine($"EAN Matching Log - {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            sb.AppendLine("=".PadRight(80, '='));
            sb.AppendLine();

            foreach (var log in allLogs)
            {
                sb.AppendLine(log);
            }

            File.WriteAllText(logFilePath, sb.ToString(), Encoding.UTF8);
        }
        catch (Exception ex)
        {
            // If we can't write the log file, at least try to log the error
            Log($"Failed to write log file: {ex.Message}", LogLevel.Error);
        }
    }

    /// <summary>
    /// Clears all buffered logs.
    /// </summary>
    public void Clear()
    {
        lock (_lockObject)
        {
            while (_logBuffer.TryDequeue(out _)) { }
        }
    }
}
