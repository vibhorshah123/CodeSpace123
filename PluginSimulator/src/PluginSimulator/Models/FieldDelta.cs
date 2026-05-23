using System;

namespace PluginSimulator.Models
{
    /// <summary>
    /// Describes a single attribute change being simulated.
    /// Used by the audit-driven popup and by ContextReconstructor.
    /// Values are kept as strings here so they survive grid editing without type juggling;
    /// ContextReconstructor parses them into typed SDK objects (OptionSetValue, EntityReference, etc.)
    /// using the explicit AttributeType — replacing the legacy ParseFieldValue heuristic.
    /// </summary>
    public class FieldDelta
    {
        public string FieldName { get; set; }

        /// <summary>Old value as a string. Empty/null means "value was empty".</summary>
        public string BeforeRaw { get; set; }

        /// <summary>New value as a string. Empty/null means "value was empty".</summary>
        public string AfterRaw { get; set; }

        /// <summary>
        /// Explicit type used by ContextReconstructor to convert raw strings to SDK types.
        /// One of: String, Int, Long, Decimal, Money, Boolean, DateTime, OptionSet, EntityReference, Guid, Null.
        /// </summary>
        public string AttributeType { get; set; } = "String";

        /// <summary>For EntityReference values, the target entity's logical name (e.g., "systemuser").</summary>
        public string EntityRefLogicalName { get; set; }

        /// <summary>Where this delta came from — for the preview's provenance icons.</summary>
        public DeltaSource Source { get; set; } = DeltaSource.Manual;

        public override string ToString() =>
            $"{FieldName} [{AttributeType}]: {BeforeRaw ?? "(null)"} → {AfterRaw ?? "(null)"} ({Source})";
    }

    public enum DeltaSource
    {
        /// <summary>Loaded from an audit history entry — accurate at that moment.</summary>
        Audit,

        /// <summary>User typed it in — hypothetical, not history.</summary>
        Manual,

        /// <summary>Pulled from the record's current Retrieve — may have drifted since the audited event.</summary>
        CurrentRetrieve
    }

    /// <summary>Known attribute types the simulator can parse. Used in the dialog combobox.</summary>
    public static class AttributeTypes
    {
        public static readonly string[] All = new[]
        {
            "String", "Int", "Long", "Decimal", "Money",
            "Boolean", "DateTime", "OptionSet", "EntityReference", "Guid", "Null"
        };

        public static string DetectFrom(object value)
        {
            if (value == null) return "Null";
            switch (value)
            {
                case Microsoft.Xrm.Sdk.OptionSetValue _:    return "OptionSet";
                case Microsoft.Xrm.Sdk.EntityReference _:   return "EntityReference";
                case Microsoft.Xrm.Sdk.Money _:             return "Money";
                case bool _:                                return "Boolean";
                case int _:                                 return "Int";
                case long _:                                return "Long";
                case decimal _:                             return "Decimal";
                case double _:                              return "Decimal";
                case DateTime _:                            return "DateTime";
                case Guid _:                                return "Guid";
                default:                                    return "String";
            }
        }
    }
}
