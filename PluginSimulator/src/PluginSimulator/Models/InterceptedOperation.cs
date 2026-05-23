using System;
using System.Collections.Generic;
using Microsoft.Xrm.Sdk;

namespace PluginSimulator.Models
{
    public enum OperationType
    {
        Create,
        Update,
        Delete,
        Execute,
        Associate,
        Disassociate
    }

    public class InterceptedOperation
    {
        public OperationType Type { get; set; }
        public string EntityName { get; set; }
        public Guid? RecordId { get; set; }
        public string RequestName { get; set; }
        public Dictionary<string, object> Fields { get; set; } = new Dictionary<string, object>();
        public Dictionary<string, object> Details { get; set; } = new Dictionary<string, object>();
        public DateTime Timestamp { get; set; } = DateTime.UtcNow;
        public Guid? ReturnedId { get; set; }
    }
}
