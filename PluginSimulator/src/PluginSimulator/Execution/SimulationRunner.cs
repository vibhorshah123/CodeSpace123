using System;
using System.Collections.Generic;
using System.Diagnostics;
using Microsoft.Xrm.Sdk;
using PluginSimulator.Models;
using PluginSimulator.Proxy;

namespace PluginSimulator.Execution
{
    /// <summary>
    /// Orchestrates the simulation: builds context, creates proxy, runs plugin.
    /// </summary>
    public static class SimulationRunner
    {
        public static SimulationResult Run(
            IPlugin plugin,
            IOrganizationService realService,
            SimulationConfig config)
        {
            var result = new SimulationResult();
            var sw = Stopwatch.StartNew();
            var copiedPdbs = new List<string>();

            try
            {
                // 1. Build execution context from real D365 data
                var context = ContextBuilder.Build(realService, config);

                // 2. Create proxy service (reads=real, writes=intercepted)
                var proxyService = new ProxyOrganizationService(realService);
                var serviceFactory = new ProxyOrganizationServiceFactory(proxyService);
                var tracingService = new ProxyTracingService();
                var logger = new ProxyLogger();

                // 3. Wire up the service provider
                var serviceProvider = new ProxyServiceProvider(context, serviceFactory, tracingService, logger);

                // 4. Optional: debug mode — copy PDB next to exe and launch debugger
                if (config.Debug)
                {
                    // Copy the plugin's PDB next to this exe so VS auto-loads symbols
                    try
                    {
                        var assemblyDir = System.IO.Path.GetDirectoryName(
                            System.IO.Path.GetFullPath(config.AssemblyPath));
                        var exeDir = System.IO.Path.GetDirectoryName(
                            System.Reflection.Assembly.GetExecutingAssembly().Location);
                        
                        foreach (var pdb in System.IO.Directory.GetFiles(assemblyDir, "*.pdb"))
                        {
                            var dest = System.IO.Path.Combine(exeDir, System.IO.Path.GetFileName(pdb));
                            System.IO.File.Copy(pdb, dest, true);
                            copiedPdbs.Add(dest);
                        }
                    }
                    catch { /* best effort */ }

                    // Debugger.Launch() opens JIT dialog → pick VS → symbols auto-loaded
                    Debugger.Launch();
                }

                // 5. Execute the plugin
                Console.WriteLine();
                Console.ForegroundColor = ConsoleColor.White;
                Console.WriteLine("  ══════════════════════════════════════════");
                Console.WriteLine("  ║       EXECUTING PLUGIN                ║");
                Console.WriteLine("  ══════════════════════════════════════════");
                Console.ResetColor();
                Console.WriteLine();

                plugin.Execute(serviceProvider);

                // 6. Collect results
                sw.Stop();
                result.Success = true;
                result.Duration = sw.Elapsed;
                result.RetrieveCount = proxyService.RetrieveCount;
                result.RetrieveMultipleCount = proxyService.RetrieveMultipleCount;
                result.InterceptedWrites = proxyService.InterceptedOperations;
                result.TraceLog = tracingService.TraceLog;
            }
            catch (InvalidPluginExecutionException ex)
            {
                sw.Stop();
                result.Success = false;
                result.Exception = ex;
                result.Duration = sw.Elapsed;
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"\n  !! Plugin threw InvalidPluginExecutionException:");
                Console.WriteLine($"     {ex.Message}");
                Console.ResetColor();
            }
            catch (Exception ex)
            {
                sw.Stop();
                result.Success = false;
                result.Exception = ex;
                result.Duration = sw.Elapsed;

                // Unwrap TargetInvocationException if present
                var inner = ex.InnerException ?? ex;
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"\n  !! Plugin threw {inner.GetType().Name}:");
                Console.WriteLine($"     {inner.Message}");
                if (inner.StackTrace != null)
                {
                    // Show first few lines of stack trace
                    var lines = inner.StackTrace.Split('\n');
                    for (int i = 0; i < Math.Min(5, lines.Length); i++)
                    {
                        Console.WriteLine($"     {lines[i].Trim()}");
                    }
                }
                Console.ResetColor();
            }

            // Cleanup copied PDBs
            foreach (var pdb in copiedPdbs)
            {
                try { System.IO.File.Delete(pdb); } catch { }
            }

            return result;
        }
    }
}
