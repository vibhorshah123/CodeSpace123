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
        public string PostImageName { get; set; } = "PostImage";

        /// <summary>
        /// Old field values loaded from audit history.
        /// These override the pre-image values fetched from the current D365 record,
        /// giving accurate "before the update" state for Update simulations.
        /// Legacy path — newer code should populate Deltas instead.
        /// </summary>
        public Dictionary<string, object> PreImageOverrides { get; set; } = new Dictionary<string, object>();

        /// <summary>
        /// Typed field deltas (Before / After values) used by the audit-driven popup.
        /// When populated, ContextBuilder delegates to ContextReconstructor and ignores
        /// ChangedFields / PreImageOverrides. Provides accurate partial Target + correct
        /// PreImage / PostImage construction per the Dataverse stage-message matrix.
        /// </summary>
        public List<FieldDelta> Deltas { get; set; } = new List<FieldDelta>();
    }
}
