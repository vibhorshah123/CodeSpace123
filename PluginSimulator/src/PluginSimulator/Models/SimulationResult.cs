using System;
using System.Collections.Generic;

namespace PluginSimulator.Models
{
    public class SimulationResult
    {
        public bool Success { get; set; }
        public Exception Exception { get; set; }
        public TimeSpan Duration { get; set; }
        public int RetrieveCount { get; set; }
        public int RetrieveMultipleCount { get; set; }
        public List<InterceptedOperation> InterceptedWrites { get; set; } = new List<InterceptedOperation>();
        public List<string> TraceLog { get; set; } = new List<string>();
    }
}
