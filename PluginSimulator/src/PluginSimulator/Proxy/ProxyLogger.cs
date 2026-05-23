using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk.PluginTelemetry;

namespace PluginSimulator.Proxy
{
    /// <summary>
    /// Fake ILogger that writes to console. Implements the full Microsoft.Xrm.Sdk.PluginTelemetry.ILogger interface.
    /// </summary>
    public class ProxyLogger : ILogger
    {
        public List<string> LogEntries { get; } = new List<string>();

        // Core log methods (string, params object[])
        public void LogCritical(string message, params object[] args) => Log("CRITICAL", message, args);
        public void LogError(string message, params object[] args) => Log("ERROR", message, args);
        public void LogWarning(string message, params object[] args) => Log("WARNING", message, args);
        public void LogInformation(string message, params object[] args) => Log("INFO", message, args);
        public void LogDebug(string message, params object[] args) => Log("DEBUG", message, args);
        public void LogTrace(string message, params object[] args) => Log("TRACE", message, args);

        // EventId overloads
        public void LogCritical(EventId eventId, Exception exception, string message, params object[] args) => Log("CRITICAL", message, args);
        public void LogCritical(EventId eventId, string message, params object[] args) => Log("CRITICAL", message, args);
        public void LogCritical(Exception exception, string message, params object[] args) => Log("CRITICAL", message, args);
        public void LogError(EventId eventId, Exception exception, string message, params object[] args) => Log("ERROR", message, args);
        public void LogError(EventId eventId, string message, params object[] args) => Log("ERROR", message, args);
        public void LogError(Exception exception, string message, params object[] args) => Log("ERROR", message, args);
        public void LogWarning(EventId eventId, Exception exception, string message, params object[] args) => Log("WARNING", message, args);
        public void LogWarning(EventId eventId, string message, params object[] args) => Log("WARNING", message, args);
        public void LogWarning(Exception exception, string message, params object[] args) => Log("WARNING", message, args);
        public void LogInformation(EventId eventId, Exception exception, string message, params object[] args) => Log("INFO", message, args);
        public void LogInformation(EventId eventId, string message, params object[] args) => Log("INFO", message, args);
        public void LogInformation(Exception exception, string message, params object[] args) => Log("INFO", message, args);
        public void LogDebug(EventId eventId, Exception exception, string message, params object[] args) => Log("DEBUG", message, args);
        public void LogDebug(EventId eventId, string message, params object[] args) => Log("DEBUG", message, args);
        public void LogDebug(Exception exception, string message, params object[] args) => Log("DEBUG", message, args);
        public void LogTrace(EventId eventId, Exception exception, string message, params object[] args) => Log("TRACE", message, args);
        public void LogTrace(EventId eventId, string message, params object[] args) => Log("TRACE", message, args);
        public void LogTrace(Exception exception, string message, params object[] args) => Log("TRACE", message, args);

        // Generic Log overloads
        public void Log(LogLevel logLevel, EventId eventId, Exception exception, string message, params object[] args) => Log(logLevel.ToString(), message, args);
        public void Log(LogLevel logLevel, EventId eventId, string message, params object[] args) => Log(logLevel.ToString(), message, args);
        public void Log(LogLevel logLevel, Exception exception, string message, params object[] args) => Log(logLevel.ToString(), message, args);
        public void Log(LogLevel logLevel, string message, params object[] args) => Log(logLevel.ToString(), message, args);
        public void Log<TState>(LogLevel logLevel, EventId eventId, TState state, Exception exception, Func<TState, Exception, string> formatter)
        {
            var message = formatter != null ? formatter(state, exception) : state?.ToString();
            Log(logLevel.ToString(), message);
        }

        // Scope, Metrics, Properties
        public IDisposable BeginScope<TState>(TState state) => new NoOpDisposable();
        public IDisposable BeginScope(string messageFormat, params object[] args) => new NoOpDisposable();
        public bool IsEnabled(LogLevel logLevel) => true;
        public void LogMetric(string metricName, long value) => Log("METRIC", $"{metricName}={value}");
        public void LogMetric(string metricName, IDictionary<string, string> metricDimensions, long value) => Log("METRIC", $"{metricName}={value}");
        public void AddCustomProperty(string propertyName, string propertyValue) { }

        // Execute
        public void Execute(string activityName, Action action, IEnumerable<KeyValuePair<string, string>> additionalCustomProperties = null)
        {
            Log("INFO", $"Executing activity: {activityName}");
            action?.Invoke();
        }

        public async Task ExecuteAsync(string activityName, Func<Task> action, IEnumerable<KeyValuePair<string, string>> additionalCustomProperties = null)
        {
            Log("INFO", $"Executing activity (async): {activityName}");
            if (action != null) await action();
        }

        private void Log(string level, string message, params object[] args)
        {
            string formatted;
            try
            {
                formatted = args != null && args.Length > 0 ? string.Format(message, args) : message;
            }
            catch { formatted = message; }

            LogEntries.Add($"[{level}] {formatted}");
            var color = level == "ERROR" || level == "CRITICAL" ? ConsoleColor.Red
                      : level == "WARNING" ? ConsoleColor.Yellow
                      : ConsoleColor.DarkCyan;
            Console.ForegroundColor = color;
            Console.WriteLine($"  LOG[{level}]: {formatted}");
            Console.ResetColor();
        }

        private class NoOpDisposable : IDisposable
        {
            public void Dispose() { }
        }
    }
}
