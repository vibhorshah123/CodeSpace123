using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.PluginTelemetry;

namespace PluginSimulator.Proxy
{
    /// <summary>
    /// Fake IServiceProvider that wires up all proxy services.
    /// The plugin calls serviceProvider.GetService(typeof(X)) to get each dependency.
    /// </summary>
    public class ProxyServiceProvider : IServiceProvider
    {
        private readonly IPluginExecutionContext _context;
        private readonly IOrganizationServiceFactory _serviceFactory;
        private readonly ITracingService _tracingService;
        private readonly ILogger _logger;

        public ProxyServiceProvider(
            IPluginExecutionContext context,
            IOrganizationServiceFactory serviceFactory,
            ITracingService tracingService,
            ILogger logger)
        {
            _context = context;
            _serviceFactory = serviceFactory;
            _tracingService = tracingService;
            _logger = logger;
        }

        public object GetService(Type serviceType)
        {
            if (serviceType == typeof(IPluginExecutionContext))
                return _context;
            if (serviceType == typeof(IOrganizationServiceFactory))
                return _serviceFactory;
            if (serviceType == typeof(ITracingService))
                return _tracingService;
            if (serviceType == typeof(ILogger))
                return _logger;

            // Return null for unknown services (plugin should handle gracefully)
            Console.ForegroundColor = ConsoleColor.DarkYellow;
            Console.WriteLine($"  ⚠ GetService requested unknown type: {serviceType.FullName}");
            Console.ResetColor();
            return null;
        }
    }
}
