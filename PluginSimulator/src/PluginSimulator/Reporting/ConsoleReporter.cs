using System;
using System.Collections.Generic;
using System.Linq;
using PluginSimulator.Models;

namespace PluginSimulator.Reporting
{
    /// <summary>
    /// Pretty-prints simulation results to the console.
    /// </summary>
    public static class ConsoleReporter
    {
        public static void Report(SimulationResult result, SimulationConfig config)
        {
            Console.WriteLine();
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine("  ╔══════════════════════════════════════════════════════════════╗");
            Console.WriteLine("  ║                    SIMULATION RESULTS                       ║");
            Console.WriteLine("  ╠══════════════════════════════════════════════════════════════╣");

            // Status
            if (result.Success)
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("  ║  Status:  ✅ COMPLETED (no exceptions)                     ║");
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("  ║  Status:  ❌ FAILED (exception thrown)                      ║");
            }

            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine($"  ║  Duration: {result.Duration.TotalMilliseconds:F0}ms                                           ║".PadRight(66) + "║");
            Console.WriteLine($"  ║  Message:  {config.MessageName} on {config.EntityLogicalName}".PadRight(65) + "║");
            Console.WriteLine($"  ║  Record:   {config.RecordId}".PadRight(65) + "║");
            Console.WriteLine("  ╠══════════════════════════════════════════════════════════════╣");
            Console.ResetColor();

            // Data Reads
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine($"  ║  DATA READS (from real D365)                                ║");
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine($"  ║    Retrieve:         {result.RetrieveCount,5} call(s)".PadRight(65) + "║");
            Console.WriteLine($"  ║    RetrieveMultiple: {result.RetrieveMultipleCount,5} call(s)".PadRight(65) + "║");
            Console.WriteLine("  ╠══════════════════════════════════════════════════════════════╣");

            // Intercepted Writes
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine($"  ║  INTERCEPTED WRITES (NOT applied to D365)                   ║");
            Console.ForegroundColor = ConsoleColor.White;

            if (result.InterceptedWrites.Count == 0)
            {
                Console.WriteLine("  ║    (none)                                                   ║");
            }
            else
            {
                // Group by type
                var grouped = result.InterceptedWrites
                    .GroupBy(o => o.Type)
                    .OrderBy(g => g.Key);

                foreach (var group in grouped)
                {
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine($"  ║                                                             ║");
                    Console.WriteLine($"  ║    {group.Key} ({group.Count()} operation(s)):".PadRight(64) + "║");
                    Console.ForegroundColor = ConsoleColor.White;

                    int opNum = 1;
                    foreach (var op in group)
                    {
                        var entity = op.EntityName ?? op.RequestName ?? "N/A";
                        var id = op.RecordId?.ToString("N").Substring(0, 8) ?? "N/A";
                        Console.WriteLine($"  ║    #{opNum}: {entity} ({id}...)".PadRight(65) + "║");

                        // Show fields for Create/Update
                        if (op.Fields.Count > 0)
                        {
                            foreach (var field in op.Fields.Take(10))
                            {
                                var val = TruncateValue(field.Value?.ToString(), 35);
                                Console.WriteLine($"  ║      {field.Key} = {val}".PadRight(65) + "║");
                            }
                            if (op.Fields.Count > 10)
                            {
                                Console.WriteLine($"  ║      ... and {op.Fields.Count - 10} more field(s)".PadRight(65) + "║");
                            }
                        }

                        // Show details for Execute operations
                        if (op.Details.Count > 0)
                        {
                            foreach (var detail in op.Details.Take(5))
                            {
                                var val = TruncateValue(detail.Value?.ToString(), 35);
                                Console.WriteLine($"  ║      {detail.Key} = {val}".PadRight(65) + "║");
                            }
                        }

                        opNum++;
                    }
                }
            }

            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine("  ╠══════════════════════════════════════════════════════════════╣");

            // Exception details
            if (!result.Success && result.Exception != null)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("  ║  EXCEPTION DETAILS                                         ║");
                Console.ForegroundColor = ConsoleColor.White;
                var exMsg = TruncateValue(result.Exception.Message, 55);
                Console.WriteLine($"  ║    Type: {result.Exception.GetType().Name}".PadRight(65) + "║");
                Console.WriteLine($"  ║    Message: {exMsg}".PadRight(65) + "║");

                var inner = result.Exception.InnerException;
                if (inner != null)
                {
                    var innerMsg = TruncateValue(inner.Message, 50);
                    Console.WriteLine($"  ║    Inner: {innerMsg}".PadRight(65) + "║");
                }

                Console.WriteLine("  ╠══════════════════════════════════════════════════════════════╣");
            }

            // Summary counts
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine("  ║  SUMMARY                                                    ║");
            Console.ForegroundColor = ConsoleColor.White;

            var creates = result.InterceptedWrites.Count(o => o.Type == OperationType.Create);
            var updates = result.InterceptedWrites.Count(o => o.Type == OperationType.Update);
            var deletes = result.InterceptedWrites.Count(o => o.Type == OperationType.Delete);
            var executes = result.InterceptedWrites.Count(o => o.Type == OperationType.Execute);

            Console.WriteLine($"  ║    Creates:    {creates,3}     Updates:   {updates,3}".PadRight(65) + "║");
            Console.WriteLine($"  ║    Deletes:    {deletes,3}     Executes:  {executes,3}".PadRight(65) + "║");
            Console.WriteLine($"  ║    Trace logs: {result.TraceLog.Count,3}                                      ║".PadRight(66) + "║");

            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine("  ╚══════════════════════════════════════════════════════════════╝");
            Console.ResetColor();
            Console.WriteLine();
        }

        private static string TruncateValue(string value, int maxLength)
        {
            if (string.IsNullOrEmpty(value)) return "(null)";
            if (value.Length <= maxLength) return value;
            return value.Substring(0, maxLength - 3) + "...";
        }
    }
}
