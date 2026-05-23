using System;
using System.Collections.Generic;
using System.Globalization;
using Microsoft.Xrm.Sdk;
using PluginSimulator.Models;

namespace PluginSimulator.Execution
{
    /// <summary>
    /// Pure reconstruction logic.
    /// Given a current Retrieve() snapshot and a set of FieldDelta entries, produces:
    ///   - Target          (only the changed fields, typed — matches real Dataverse Update semantics)
    ///   - PreImage        (full snapshot with Before values overlaid)
    ///   - PostImage       (PreImage + Target overlay) — gated by stage/message matrix per MS Learn
    ///   - Provenance map  (per-attribute source: Audit / Manual / CurrentRetrieve)
    /// </summary>
    public static class ContextReconstructor
    {
        public class ReconstructedContext
        {
            public Entity Target { get; set; }
            public Entity PreImage { get; set; }
            public Entity PostImage { get; set; } // null when not applicable
            public Dictionary<string, DeltaSource> PreImageProvenance { get; set; } = new Dictionary<string, DeltaSource>();
        }

        public static ReconstructedContext Reconstruct(
            string entityLogicalName,
            Guid recordId,
            Entity currentRecord,
            IList<FieldDelta> deltas,
            string messageName,
            int stage)
        {
            if (deltas == null) deltas = new List<FieldDelta>();

            var result = new ReconstructedContext
            {
                Target = BuildTarget(entityLogicalName, recordId, deltas, messageName, currentRecord),
                PreImage = BuildPreImage(entityLogicalName, recordId, currentRecord, deltas, messageName, out var provenance)
            };
            result.PreImageProvenance = provenance;

            result.PostImage = BuildPostImage(
                result.PreImage,
                result.Target,
                messageName,
                stage,
                entityLogicalName,
                recordId);

            return result;
        }

        private static Entity BuildTarget(
            string entityLogicalName,
            Guid recordId,
            IList<FieldDelta> deltas,
            string messageName,
            Entity currentRecord)
        {
            if (messageName.Equals("Delete", StringComparison.OrdinalIgnoreCase))
            {
                // Delete uses EntityReference as Target — but we return Entity here for uniformity;
                // ContextBuilder converts to EntityReference when stitching into InputParameters.
                return new Entity(entityLogicalName, recordId);
            }

            if (messageName.Equals("Create", StringComparison.OrdinalIgnoreCase))
            {
                // Create: Target is the full record being created.
                // Best heuristic when we have a current Retrieve: use After values overlaid on the current record.
                var target = new Entity(entityLogicalName, recordId);
                if (currentRecord != null)
                {
                    foreach (var attr in currentRecord.Attributes)
                        target[attr.Key] = attr.Value;
                }
                foreach (var d in deltas)
                {
                    var v = ParseValue(d.AfterRaw, d.AttributeType, d.EntityRefLogicalName);
                    target[d.FieldName] = v;
                }
                return target;
            }

            // Update (and any other message): Target = ONLY the changed fields, typed.
            // This matches the MS Learn rule: "Only include columns with changed values in update operations."
            var updateTarget = new Entity(entityLogicalName, recordId);
            foreach (var d in deltas)
            {
                if (string.IsNullOrWhiteSpace(d.FieldName)) continue;
                var v = ParseValue(d.AfterRaw, d.AttributeType, d.EntityRefLogicalName);
                updateTarget[d.FieldName] = v;
            }
            return updateTarget;
        }

        private static Entity BuildPreImage(
            string entityLogicalName,
            Guid recordId,
            Entity currentRecord,
            IList<FieldDelta> deltas,
            string messageName,
            out Dictionary<string, DeltaSource> provenance)
        {
            provenance = new Dictionary<string, DeltaSource>();

            // Create has no PreImage per MS Learn ("the table doesn't exist yet").
            if (messageName.Equals("Create", StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }

            // Start with a deep-ish clone of currentRecord so plugin mutations don't pollute the live cache.
            var preImage = new Entity(entityLogicalName, recordId);
            if (currentRecord != null)
            {
                foreach (var attr in currentRecord.Attributes)
                {
                    preImage[attr.Key] = attr.Value;
                    provenance[attr.Key] = DeltaSource.CurrentRetrieve;
                }
            }

            // Overlay BeforeRaw values from deltas — these are what audit told us the field was at that moment.
            foreach (var d in deltas)
            {
                if (string.IsNullOrWhiteSpace(d.FieldName)) continue;
                var v = ParseValue(d.BeforeRaw, d.AttributeType, d.EntityRefLogicalName);
                preImage[d.FieldName] = v;
                provenance[d.FieldName] = d.Source;
            }

            return preImage;
        }

        /// <summary>
        /// Build PostImage following the MS Learn availability matrix:
        ///   Create + PostOp    → PostImage = the created record (Target)
        ///   Update + PostOp    → PostImage = PreImage + Target overlay
        ///   Delete             → no PostImage
        ///   Any non-PostOp     → no PostImage
        /// </summary>
        private static Entity BuildPostImage(
            Entity preImage,
            Entity target,
            string messageName,
            int stage,
            string entityLogicalName,
            Guid recordId)
        {
            if (stage != 40) return null; // PostImage only at PostOperation
            if (messageName.Equals("Delete", StringComparison.OrdinalIgnoreCase)) return null;

            if (messageName.Equals("Create", StringComparison.OrdinalIgnoreCase))
            {
                // Post-Create image = the freshly-created record. Target carries it for Create.
                if (target == null) return null;
                var post = new Entity(entityLogicalName, recordId);
                foreach (var attr in target.Attributes)
                    post[attr.Key] = attr.Value;
                return post;
            }

            // Update: PostImage = PreImage clone + Target overlay
            if (preImage == null) return null;
            var postImage = new Entity(entityLogicalName, recordId);
            foreach (var attr in preImage.Attributes)
                postImage[attr.Key] = attr.Value;
            if (target != null)
            {
                foreach (var attr in target.Attributes)
                    postImage[attr.Key] = attr.Value;
            }
            return postImage;
        }

        /// <summary>
        /// Parse a string value into the typed SDK object based on the explicit attribute type.
        /// Replaces the legacy ParseFieldValue heuristic in ContextBuilder that silently wrapped
        /// ints as OptionSetValue and decimals as Money.
        /// </summary>
        public static object ParseValue(string raw, string type, string entityRefLogicalName)
        {
            if (string.IsNullOrEmpty(raw) || type == "Null") return null;
            var ci = CultureInfo.InvariantCulture;

            switch (type)
            {
                case "String":
                    return raw;

                case "Int":
                    if (int.TryParse(raw, NumberStyles.Integer, ci, out var iv)) return iv;
                    throw new FormatException($"Cannot parse '{raw}' as Int");

                case "Long":
                    if (long.TryParse(raw, NumberStyles.Integer, ci, out var lv)) return lv;
                    throw new FormatException($"Cannot parse '{raw}' as Long");

                case "Decimal":
                    if (decimal.TryParse(raw, NumberStyles.Number, ci, out var dv)) return dv;
                    throw new FormatException($"Cannot parse '{raw}' as Decimal");

                case "Money":
                    if (decimal.TryParse(raw, NumberStyles.Number, ci, out var mv)) return new Money(mv);
                    throw new FormatException($"Cannot parse '{raw}' as Money");

                case "Boolean":
                    if (bool.TryParse(raw, out var bv)) return bv;
                    if (raw == "1") return true;
                    if (raw == "0") return false;
                    throw new FormatException($"Cannot parse '{raw}' as Boolean");

                case "DateTime":
                    if (DateTime.TryParse(raw, ci, DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal, out var dtv))
                        return dtv;
                    throw new FormatException($"Cannot parse '{raw}' as DateTime");

                case "OptionSet":
                    if (int.TryParse(raw, NumberStyles.Integer, ci, out var ov)) return new OptionSetValue(ov);
                    throw new FormatException($"Cannot parse '{raw}' as OptionSet (must be integer)");

                case "EntityReference":
                    if (!Guid.TryParse(raw, out var refId))
                        throw new FormatException($"Cannot parse '{raw}' as EntityReference Guid");
                    if (string.IsNullOrWhiteSpace(entityRefLogicalName))
                        throw new FormatException($"EntityReference for '{raw}' is missing the target entity logical name");
                    return new EntityReference(entityRefLogicalName, refId);

                case "Guid":
                    if (Guid.TryParse(raw, out var gv)) return gv;
                    throw new FormatException($"Cannot parse '{raw}' as Guid");

                default:
                    throw new NotSupportedException($"Unknown attribute type '{type}'");
            }
        }

        /// <summary>
        /// Inverse of ParseValue — format a typed SDK value into a string for the grid.
        /// </summary>
        public static string FormatValue(object value)
        {
            if (value == null) return "";
            switch (value)
            {
                case OptionSetValue osv:    return osv.Value.ToString(CultureInfo.InvariantCulture);
                case EntityReference er:    return er.Id.ToString();
                case Money m:               return m.Value.ToString(CultureInfo.InvariantCulture);
                case bool b:                return b ? "true" : "false";
                case DateTime dt:           return dt.ToUniversalTime().ToString("o", CultureInfo.InvariantCulture);
                case AliasedValue av:       return FormatValue(av.Value);
                default:                    return Convert.ToString(value, CultureInfo.InvariantCulture);
            }
        }
    }
}
