using System.Collections.Generic;

namespace PluginSimulator.Models
{
    public class SimulationConfig
    {
        public string EnvironmentUrl { get; set; }
        public string EntityLogicalName { get; set; }
        public string RecordId { get; set; }
        public string MessageName { get; set; } = "Create";
        public int Stage { get; set; } = 40; // Post-Operation
        public int Depth { get; set; } = 1;
        public string PluginTypeName { get; set; }
        public string AssemblyPath { get; set; }
        public bool Debug { get; set; }
        public string ConnectionString { get; set; }
        public Dictionary<string, string> ChangedFields { get; set; } = new Dictionary<string, string>();
        public string PreImageName { get; set; } = "PreImage";

        /// <summary>
        /// Old field values loaded from audit history.
        /// These override the pre-image values fetched from the current D365 record,
        /// giving accurate "before the update" state for Update simulations.
        /// </summary>
        public Dictionary<string, object> PreImageOverrides { get; set; } = new Dictionary<string, object>();
    }
}
