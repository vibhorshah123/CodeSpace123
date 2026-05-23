using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using PluginSimulator.Models;

namespace PluginSimulator.Execution
{
    /// <summary>
    /// Builds a fake IPluginExecutionContext from a real D365 record.
    /// For Create: fetches the full record and uses it as Target.
    /// For Update: fetches the current record as PreImage, user specifies changed fields as Target.
    /// </summary>
    public static class ContextBuilder
    {
        public static Proxy.ProxyExecutionContext Build(
            IOrganizationService realService,
            SimulationConfig config)
        {
            var recordId = Guid.Parse(config.RecordId);

            Console.WriteLine();
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine("  Building execution context...");
            Console.ResetColor();

            // Fetch the full record from D365
            var fullRecord = realService.Retrieve(
                config.EntityLogicalName,
                recordId,
                new ColumnSet(true));

            Console.WriteLine($"  Fetched {config.EntityLogicalName} ({recordId})");
            Console.WriteLine($"  Attributes: {fullRecord.Attributes.Count}");

            var context = new Proxy.ProxyExecutionContext
            {
                MessageName = config.MessageName,
                Stage = config.Stage,
                Depth = config.Depth,
                PrimaryEntityName = config.EntityLogicalName,
                PrimaryEntityId = recordId,
                OrganizationName = "SimulatedOrg"
            };

            if (config.Deltas != null && config.Deltas.Count > 0)
            {
                BuildFromDeltas(context, config, fullRecord, recordId);
            }
            else if (config.MessageName.Equals("Create", StringComparison.OrdinalIgnoreCase))
            {
                // For Create: the full record IS the target
                context.InputParameters["Target"] = fullRecord;
                Console.WriteLine($"  Mode: Create → full record used as Target");
            }
            else if (config.MessageName.Equals("Update", StringComparison.OrdinalIgnoreCase))
            {
                // For Update: build Target with only changed fields, PreImage = current full record
                var targetEntity = new Entity(config.EntityLogicalName, recordId);

                if (config.ChangedFields != null && config.ChangedFields.Count > 0)
                {
                    // User specified which fields changed — copy their CURRENT values from the real record
                    foreach (var field in config.ChangedFields)
                    {
                        if (fullRecord.Attributes.ContainsKey(field.Key))
                        {
                            targetEntity[field.Key] = fullRecord[field.Key];
                        }
                        else
                        {
                            // Field specified but doesn't exist on record — set from user input
                            targetEntity[field.Key] = ParseFieldValue(field.Value);
                        }
                    }
                    Console.WriteLine($"  Mode: Update → {targetEntity.Attributes.Count} changed field(s) in Target");
                }
                else
                {
                    // No specific fields — use full record as target (simulates full-record update)
                    foreach (var attr in fullRecord.Attributes)
                    {
                        targetEntity[attr.Key] = attr.Value;
                    }
                    Console.WriteLine($"  Mode: Update → full record used as Target (no specific changed fields)");
                }

                context.InputParameters["Target"] = targetEntity;

                // Build pre-image from the full record using configurable name
                var preImage = new Entity(config.EntityLogicalName, recordId);
                foreach (var attr in fullRecord.Attributes)
                {
                    preImage[attr.Key] = attr.Value;
                }
                context.PreEntityImages[config.PreImageName] = preImage;
                Console.WriteLine($"  PreImage '{config.PreImageName}': {preImage.Attributes.Count} attribute(s)");

                // Apply audit-sourced old values on top so pre-image reflects real state before the update
                ApplyPreImageOverrides(preImage, config);
            }
            else if (config.MessageName.Equals("Delete", StringComparison.OrdinalIgnoreCase))
            {
                // For Delete: Target is an EntityReference; PreImage = full record before deletion
                context.InputParameters["Target"] = fullRecord.ToEntityReference();
                var preImage = new Entity(config.EntityLogicalName, recordId);
                foreach (var attr in fullRecord.Attributes)
                    preImage[attr.Key] = attr.Value;
                context.PreEntityImages[config.PreImageName] = preImage;
                Console.WriteLine($"  Mode: Delete → EntityReference as Target");
                Console.WriteLine($"  PreImage '{config.PreImageName}': {preImage.Attributes.Count} attribute(s)");
                ApplyPreImageOverrides(preImage, config);
            }
            else
            {
                // Custom message: put full record as Target
                context.InputParameters["Target"] = fullRecord;
                context.InputParameters["ProcessName"] = config.MessageName;
                Console.WriteLine($"  Mode: Custom message '{config.MessageName}'");
            }

            // Set user context from the real service
            try
            {
                var whoAmI = (Microsoft.Crm.Sdk.Messages.WhoAmIResponse)realService.Execute(
                    new Microsoft.Crm.Sdk.Messages.WhoAmIRequest());
                context.UserId = whoAmI.UserId;
                context.InitiatingUserId = whoAmI.UserId;
                context.BusinessUnitId = whoAmI.BusinessUnitId;
                context.OrganizationId = whoAmI.OrganizationId;
                Console.WriteLine($"  User: {whoAmI.UserId}");
            }
            catch
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine("  ⚠ Could not get WhoAmI — using default user context");
                Console.ResetColor();
            }

            return context;
        }

        /// <summary>
        /// Audit-driven path: reconstructs Target / PreImage / PostImage from FieldDelta entries
        /// using ContextReconstructor (which enforces the MS Learn stage-message matrix).
        /// </summary>
        private static void BuildFromDeltas(
            Proxy.ProxyExecutionContext context,
            SimulationConfig config,
            Entity fullRecord,
            Guid recordId)
        {
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine($"  Mode: Audit-driven reconstruction ({config.Deltas.Count} delta(s))");
            Console.ResetColor();

            var rebuilt = ContextReconstructor.Reconstruct(
                config.EntityLogicalName,
                recordId,
                fullRecord,
                config.Deltas,
                config.MessageName,
                config.Stage);

            // Target
            if (config.MessageName.Equals("Delete", StringComparison.OrdinalIgnoreCase))
            {
                // Delete uses EntityReference as Target
                context.InputParameters["Target"] = new EntityReference(config.EntityLogicalName, recordId);
            }
            else
            {
                context.InputParameters["Target"] = rebuilt.Target;
            }
            Console.WriteLine($"  Target: {(rebuilt.Target?.Attributes.Count ?? 0)} attribute(s)");

            // PreImage (skip for Create per MS Learn)
            if (rebuilt.PreImage != null)
            {
                context.PreEntityImages[config.PreImageName] = rebuilt.PreImage;
                Console.WriteLine($"  PreImage '{config.PreImageName}': {rebuilt.PreImage.Attributes.Count} attribute(s)");
            }

            // PostImage (only PostOp + Create/Update per MS Learn)
            if (rebuilt.PostImage != null)
            {
                context.PostEntityImages[config.PostImageName] = rebuilt.PostImage;
                Console.WriteLine($"  PostImage '{config.PostImageName}': {rebuilt.PostImage.Attributes.Count} attribute(s)");
            }
            else if (config.Stage == 40 &&
                     !config.MessageName.Equals("Delete", StringComparison.OrdinalIgnoreCase))
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine("  PostImage: not built (stage/message combination not eligible)");
                Console.ResetColor();
            }
        }

        private static void ApplyPreImageOverrides(Entity preImage, SimulationConfig config)
        {
            if (config.PreImageOverrides == null || config.PreImageOverrides.Count == 0) return;
            foreach (var kvp in config.PreImageOverrides)
                preImage[kvp.Key] = kvp.Value;
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine($"  Applied {config.PreImageOverrides.Count} pre-image override(s) from audit history");
            Console.ResetColor();
        }

        private static object ParseFieldValue(string value)        {
            if (string.IsNullOrEmpty(value)) return null;
            if (int.TryParse(value, out int intVal)) return new OptionSetValue(intVal);
            if (Guid.TryParse(value, out Guid guidVal)) return guidVal;
            if (DateTime.TryParse(value, out DateTime dateVal)) return dateVal;
            if (bool.TryParse(value, out bool boolVal)) return boolVal;
            if (decimal.TryParse(value, out decimal decVal)) return new Money(decVal);
            return value;
        }
    }
}
