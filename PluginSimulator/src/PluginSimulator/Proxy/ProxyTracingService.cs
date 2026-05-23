using System;
using System.Collections.Generic;
using Microsoft.Xrm.Sdk;

namespace PluginSimulator.Proxy
{
    /// <summary>
    /// Captures trace output from the plugin into a list for reporting.
    /// </summary>
    public class ProxyTracingService : ITracingService
    {
        public List<string> TraceLog { get; } = new List<string>();

        public void Trace(string format, params object[] args)
        {
            string message;
            try
            {
                message = args != null && args.Length > 0
                    ? string.Format(format, args)
                    : format;
            }
            catch
            {
                message = format;
            }

            TraceLog.Add($"[{DateTime.UtcNow:HH:mm:ss.fff}] {message}");

            Console.ForegroundColor = ConsoleColor.DarkGray;
            Console.WriteLine($"  TRACE: {message}");
            Console.ResetColor();
        }
    }
}
