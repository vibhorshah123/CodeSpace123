using System;
using System.Collections.Generic;
using Microsoft.Xrm.Sdk;

namespace PluginSimulator.Proxy
{
    /// <summary>
    /// Fake IPluginExecutionContext built from real D365 record data.
    /// </summary>
    public class ProxyExecutionContext : IPluginExecutionContext
    {
        public int Stage { get; set; } = 40;
        public IPluginExecutionContext ParentContext { get; set; }
        public int Mode { get; set; } = 0; // Synchronous
        public int IsolationMode { get; set; } = 2; // Sandbox
        public int Depth { get; set; } = 1;
        public string MessageName { get; set; } = "Create";
        public string PrimaryEntityName { get; set; }
        public Guid? RequestId { get; set; } = Guid.NewGuid();
        public string SecondaryEntityName { get; set; } = "";
        public ParameterCollection InputParameters { get; set; } = new ParameterCollection();
        public ParameterCollection OutputParameters { get; set; } = new ParameterCollection();
        public ParameterCollection SharedVariables { get; set; } = new ParameterCollection();
        public EntityImageCollection PreEntityImages { get; set; } = new EntityImageCollection();
        public EntityImageCollection PostEntityImages { get; set; } = new EntityImageCollection();
        public Guid UserId { get; set; } = Guid.NewGuid();
        public Guid InitiatingUserId { get; set; } = Guid.NewGuid();
        public Guid BusinessUnitId { get; set; } = Guid.NewGuid();
        public Guid OrganizationId { get; set; } = Guid.NewGuid();
        public string OrganizationName { get; set; } = "SimulatedOrg";
        public Guid PrimaryEntityId { get; set; }
        public Guid CorrelationId { get; set; } = Guid.NewGuid();
        public bool IsExecutingOffline { get; set; }
        public bool IsOfflinePlayback { get; set; }
        public bool IsInTransaction { get; set; } = true;
        public Guid OperationId { get; set; } = Guid.NewGuid();
        public DateTime OperationCreatedOn { get; set; } = DateTime.UtcNow;
        public EntityReference OwningExtension { get; set; }
    }
}
