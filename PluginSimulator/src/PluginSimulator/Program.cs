using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using PluginSimulator.Authentication;
using PluginSimulator.Execution;
using PluginSimulator.Models;
using PluginSimulator.Reporting;
using PluginSimulator.UI;

namespace PluginSimulator
{
    class Program
    {
        [STAThread]
        static int Main(string[] args)
        {
            // Resolve DLLs from 'lib' subfolder next to this exe
            var exeDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            var libDir = Path.Combine(exeDir, "lib");
            AppDomain.CurrentDomain.AssemblyResolve += (sender, resolveArgs) =>
            {
                var dllName = new System.Reflection.AssemblyName(resolveArgs.Name).Name + ".dll";
                var dllPath = Path.Combine(libDir, dllName);
                if (File.Exists(dllPath))
                    return System.Reflection.Assembly.LoadFrom(dllPath);
                return null;
            };

            return Run(args);
        }

        public static int Run(string[] args)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;

            // No args → launch WinForms UI
            if (args.Length == 0)
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.ThreadException += (s, e) =>
                    MessageBox.Show(e.Exception.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                AppDomain.CurrentDomain.UnhandledException += (s, e) =>
                    MessageBox.Show(e.ExceptionObject.ToString(), "Fatal Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Run(new MainForm());
                return 0;
            }

            // CLI mode (backward compatible)
            PrintBanner();

            if (args.Contains("--help") || args.Contains("-h"))
            {
                PrintUsage();
                return 0;
            }

            try
            {
                var config = ParseArgs(args);
                ValidateConfig(config);

                // Step 1: Connect to D365
                Console.ForegroundColor = ConsoleColor.Cyan;
                Console.WriteLine("\n  [1/4] Connecting to D365...");
                Console.ResetColor();
                var realService = AuthManager.Connect(config.EnvironmentUrl, config.ConnectionString);

                // Step 2: Load plugin assembly
                Console.ForegroundColor = ConsoleColor.Cyan;
                Console.WriteLine("\n  [2/4] Loading plugin assembly...");
                Console.ResetColor();
                var plugin = PluginLoader.Load(config.AssemblyPath, config.PluginTypeName);

                // Step 3: Run simulation
                Console.ForegroundColor = ConsoleColor.Cyan;
                Console.WriteLine("\n  [3/4] Running simulation...");
                Console.ResetColor();
                var result = SimulationRunner.Run(plugin, realService, config);

                // Step 4: Report results
                Console.ForegroundColor = ConsoleColor.Cyan;
                Console.WriteLine("\n  [4/4] Results:");
                Console.ResetColor();
                ConsoleReporter.Report(result, config);

                return result.Success ? 0 : 1;
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"\n  FATAL ERROR: {ex.Message}");
                if (ex.InnerException != null)
                    Console.WriteLine($"  Inner: {ex.InnerException.Message}");
                Console.ResetColor();
                return 2;
            }
        }

        static SimulationConfig ParseArgs(string[] args)
        {
            var config = new SimulationConfig();

            for (int i = 0; i < args.Length; i++)
            {
                var arg = args[i].ToLowerInvariant();
                var next = i + 1 < args.Length ? args[i + 1] : null;

                switch (arg)
                {
                    case "--env":
                    case "-e":
                        config.EnvironmentUrl = next;
                        i++;
                        break;

                    case "--entity":
                        config.EntityLogicalName = next;
                        i++;
                        break;

                    case "--id":
                        config.RecordId = next;
                        i++;
                        break;

                    case "--message":
                    case "-m":
                        config.MessageName = next;
                        i++;
                        break;

                    case "--stage":
                        config.Stage = int.Parse(next);
                        i++;
                        break;

                    case "--depth":
                        config.Depth = int.Parse(next);
                        i++;
                        break;

                    case "--plugin":
                    case "-p":
                        config.PluginTypeName = next;
                        i++;
                        break;

                    case "--assembly":
                    case "-a":
                        config.AssemblyPath = next;
                        i++;
                        break;

                    case "--connection-string":
                    case "-cs":
                        config.ConnectionString = next;
                        i++;
                        break;

                    case "--debug":
                        config.Debug = true;
                        break;

                    case "--changed":
                        // Format: field1=value1,field2=value2
                        if (next != null)
                        {
                            foreach (var pair in next.Split(','))
                            {
                                var parts = pair.Split(new[] { '=' }, 2);
                                if (parts.Length == 2)
                                    config.ChangedFields[parts[0].Trim()] = parts[1].Trim();
                            }
                            i++;
                        }
                        break;
                }
            }

            return config;
        }

        static void ValidateConfig(SimulationConfig config)
        {
            var errors = new List<string>();

            if (string.IsNullOrEmpty(config.EnvironmentUrl))
                errors.Add("--env is required (e.g., https://mashppe.crm.dynamics.com)");
            if (string.IsNullOrEmpty(config.EntityLogicalName))
                errors.Add("--entity is required (e.g., mash_servicetask)");
            if (string.IsNullOrEmpty(config.RecordId))
                errors.Add("--id is required (record GUID)");
            if (string.IsNullOrEmpty(config.AssemblyPath))
                errors.Add("--assembly is required (path to plugin DLL)");

            if (errors.Count > 0)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("\n  Missing required arguments:");
                foreach (var e in errors)
                    Console.WriteLine($"    • {e}");
                Console.ResetColor();
                Console.WriteLine();
                PrintUsage();
                throw new ArgumentException("Missing required arguments.");
            }
        }

        static void PrintBanner()
        {
            Console.ForegroundColor = ConsoleColor.Magenta;
            Console.WriteLine(@"
  ╔══════════════════════════════════════════════════════════╗
  ║          D365 Plugin Simulator v1.0                     ║
  ║          Reads=Real D365 | Writes=Intercepted           ║
  ╚══════════════════════════════════════════════════════════╝");
            Console.ResetColor();
        }

        static void PrintUsage()
        {
            Console.WriteLine(@"
  USAGE:
    PluginSimulator.exe [options]

  REQUIRED:
    --env, -e       D365 environment URL
                    e.g., https://mashppe.crm.dynamics.com

    --entity        Entity logical name
                    e.g., mash_servicetask

    --id            Record GUID to simulate against
                    e.g., aaaa-bbbb-cccc-dddd-eeee

    --assembly, -a  Path to plugin DLL
                    e.g., D:\MaSH\Plugins\bin\Debug\MASHvNext.ServiceTask.dll

  OPTIONAL:
    --plugin, -p    Full type name of plugin class
                    e.g., MASHvNext.ServiceTask.ServiceTask
                    (auto-detected if only one IPlugin in assembly)

    --message, -m   Message name (default: Create)
                    e.g., Create, Update, Delete

    --stage         Plugin stage (default: 40 = Post-Operation)
                    10=PreValidation, 20=PreOperation, 40=PostOperation

    --depth         Execution depth (default: 1)

    --changed       For Update: comma-separated changed fields
                    e.g., --changed ""mash_activestage=Build,mash_status=100000002""

    --debug         Launch Visual Studio debugger before executing plugin

    --connection-string, -cs
                    Full CRM connection string (overrides --env)

  EXAMPLES:
    # Simulate Create of a service task
    PluginSimulator.exe --env https://mashppe.crm.dynamics.com ^
      --entity mash_servicetask ^
      --id 12345678-1234-1234-1234-123456789012 ^
      --assembly ""D:\MaSH On Dynamics\Plugins\MASHvNext.ServiceTask\bin\Debug\MASHvNext.ServiceTask.dll"" ^
      --message Create

    # Simulate Update with debugger
    PluginSimulator.exe --env https://mashppe.crm.dynamics.com ^
      --entity mash_servicetask ^
      --id 12345678-1234-1234-1234-123456789012 ^
      --assembly ""D:\MaSH On Dynamics\Plugins\MASHvNext.ServiceTask\bin\Debug\MASHvNext.ServiceTask.dll"" ^
      --message Update ^
      --changed ""mash_activestage=Build"" ^
      --debug
");
        }
    }
}
