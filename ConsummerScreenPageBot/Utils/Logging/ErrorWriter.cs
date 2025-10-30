using System;
using System.IO;

namespace ConsummerScreenPageBot
{
    public static class ErrorWriter
    {
        public static void WriteLog(string logPath, string category, string message)
        {
            try
            {
                if (!Directory.Exists(logPath))
                {
                    Directory.CreateDirectory(logPath);
                }
                string logFile = Path.Combine(logPath, $"error_{DateTime.Now:yyyyMMdd}.log");
                string logMessage = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [{category}] {message}\n";
                File.AppendAllText(logFile, logMessage);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to write log: {ex.Message}");
            }
        }
    }
}


