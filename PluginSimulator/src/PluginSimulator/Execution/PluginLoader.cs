using System;
using System.IO;
using System.Linq;
using System.Reflection;
using Microsoft.Xrm.Sdk;

namespace PluginSimulator.Execution
{
    /// <summary>
    /// Loads a plugin DLL via reflection and instantiates the IPlugin type.
    /// </summary>
    public static class PluginLoader
    {
        public static IPlugin Load(string assemblyPath, string pluginTypeName)
        {
            if (!File.Exists(assemblyPath))
                throw new FileNotFoundException($"Plugin assembly not found: {assemblyPath}");

            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine($"  Loading assembly: {Path.GetFileName(assemblyPath)}");

            // Add the assembly's directory to the resolution path
            var assemblyDir = Path.GetDirectoryName(Path.GetFullPath(assemblyPath));
            AppDomain.CurrentDomain.AssemblyResolve += (sender, args) =>
            {
                var name = new AssemblyName(args.Name).Name + ".dll";
                var path = Path.Combine(assemblyDir, name);
                if (File.Exists(path))
                    return Assembly.LoadFrom(path);
                return null;
            };

            var assembly = Assembly.LoadFrom(assemblyPath);
            Console.WriteLine($"  Assembly loaded: {assembly.GetName().Name} v{assembly.GetName().Version}");

            // Find the plugin type
            Type pluginType = null;

            if (!string.IsNullOrEmpty(pluginTypeName))
            {
                pluginType = assembly.GetType(pluginTypeName);
                if (pluginType == null)
                {
                    // Try partial match
                    pluginType = assembly.GetTypes()
                        .FirstOrDefault(t => t.FullName.EndsWith(pluginTypeName, StringComparison.OrdinalIgnoreCase)
                                          || t.Name.Equals(pluginTypeName, StringComparison.OrdinalIgnoreCase));
                }
            }

            if (pluginType == null)
            {
                // Auto-discover: find all IPlugin implementations
                var pluginTypes = assembly.GetTypes()
                    .Where(t => typeof(IPlugin).IsAssignableFrom(t) && !t.IsAbstract && !t.IsInterface)
                    .ToList();

                if (pluginTypes.Count == 0)
                    throw new Exception($"No IPlugin implementations found in {assemblyPath}");

                if (pluginTypes.Count == 1)
                {
                    pluginType = pluginTypes[0];
                    Console.WriteLine($"  Auto-discovered plugin: {pluginType.FullName}");
                }
                else
                {
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine($"  Multiple IPlugin types found:");
                    for (int i = 0; i < pluginTypes.Count; i++)
                    {
                        Console.WriteLine($"    [{i + 1}] {pluginTypes[i].FullName}");
                    }
                    Console.ResetColor();
                    throw new Exception("Multiple IPlugin types found. Specify --plugin <FullTypeName>");
                }
            }

            Console.WriteLine($"  Plugin type: {pluginType.FullName}");

            // Instantiate the plugin
            IPlugin plugin;
            try
            {
                plugin = (IPlugin)Activator.CreateInstance(pluginType);
            }
            catch (Exception ex)
            {
                // Some plugins require constructor parameters (unsecure/secure config)
                try
                {
                    plugin = (IPlugin)Activator.CreateInstance(pluginType, string.Empty, string.Empty);
                }
                catch
                {
                    try
                    {
                        plugin = (IPlugin)Activator.CreateInstance(pluginType, string.Empty);
                    }
                    catch
                    {
                        throw new Exception($"Cannot instantiate {pluginType.FullName}. " +
                            $"Ensure it has a parameterless constructor. Original error: {ex.Message}");
                    }
                }
            }

            Console.WriteLine($"  Plugin instantiated successfully ✓");
            Console.ResetColor();

            return plugin;
        }
    }
}
