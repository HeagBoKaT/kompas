using System;
using System.IO;

namespace GUI;

public class Logger
{
    private string _logFilePath;

    public Logger(string logFilePath)
    {
        _logFilePath = logFilePath;
    }

    public void Log(string message, string level = "INFO")
    {
        var logEntry = $"[{DateTime.Now:U}] [{level}] {message}{Environment.NewLine}";
        File.AppendAllText(_logFilePath, logEntry);
    }

    public void Info(string message) => Log(message, "INFO");
    public void Warning(string message) => Log(message, "WARNING");
    public void Error(string message) => Log(message, "ERROR");

}