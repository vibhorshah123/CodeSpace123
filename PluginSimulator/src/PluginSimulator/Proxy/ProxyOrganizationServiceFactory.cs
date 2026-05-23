using System;
using Microsoft.Xrm.Sdk;

namespace PluginSimulator.Proxy
{
    /// <summary>
    /// Fake IOrganizationServiceFactory that always returns the proxy service.
    /// </summary>
    public class ProxyOrganizationServiceFactory : IOrganizationServiceFactory
    {
        private readonly IOrganizationService _service;

        public ProxyOrganizationServiceFactory(IOrganizationService service)
        {
            _service = service;
        }

        public IOrganizationService CreateOrganizationService(Guid? userId)
        {
            return _service;
        }
    }
}
