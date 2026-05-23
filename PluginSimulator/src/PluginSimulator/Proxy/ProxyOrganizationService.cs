using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using PluginSimulator.Models;

namespace PluginSimulator.Proxy
{
    /// <summary>
    /// Wraps a real IOrganizationService. 
    /// READ operations pass through to real D365.
    /// WRITE operations are intercepted, logged, and stored in a dirty cache.
    /// Subsequent reads merge dirty cache values so the plugin sees its own writes.
    /// </summary>
    public class ProxyOrganizationService : IOrganizationService
    {
        private readonly IOrganizationService _realService;
        public List<InterceptedOperation> InterceptedOperations { get; } = new List<InterceptedOperation>();
        public int RetrieveCount { get; private set; }
        public int RetrieveMultipleCount { get; private set; }

        // Dirty cache: key = "entityName|recordId", value = dictionary of updated fields
        private readonly Dictionary<string, Dictionary<string, object>> _pendingUpdates = new Dictionary<string, Dictionary<string, object>>();

        // Created records cache: key = "entityName|newId"
        private readonly Dictionary<string, Entity> _createdRecords = new Dictionary<string, Entity>();

        // Track deleted record IDs
        private readonly HashSet<string> _deletedRecords = new HashSet<string>();

        // Set of request types that should be intercepted (writes)
        private static readonly HashSet<string> WriteRequestTypes = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            nameof(CreateRequest), nameof(UpdateRequest), nameof(DeleteRequest),
            nameof(SendEmailRequest), nameof(GrantAccessRequest), nameof(RevokeAccessRequest),
            nameof(ModifyAccessRequest), nameof(AssignRequest), nameof(SetStateRequest),
            "SendEmailFromTemplateRequest", "AddToQueueRequest", "RemoveFromQueueRequest",
            "QualifyLeadRequest", "WinOpportunityRequest", "LoseOpportunityRequest",
            "CloseIncidentRequest", "ExecuteMultipleRequest"
        };

        public ProxyOrganizationService(IOrganizationService realService)
        {
            _realService = realService ?? throw new ArgumentNullException(nameof(realService));
        }

        public Guid Create(Entity entity)
        {
            var newId = Guid.NewGuid();
            var op = new InterceptedOperation
            {
                Type = OperationType.Create,
                EntityName = entity.LogicalName,
                RecordId = newId,
                ReturnedId = newId
            };

            foreach (var attr in entity.Attributes)
            {
                op.Fields[attr.Key] = FormatValue(attr.Value);
            }

            InterceptedOperations.Add(op);

            // Store in created records cache
            entity.Id = newId;
            _createdRecords[$"{entity.LogicalName}|{newId}"] = entity;

            LogIntercepted("CREATE", entity.LogicalName, newId, op.Fields.Count);
            return newId;
        }

        public void Update(Entity entity)
        {
            var op = new InterceptedOperation
            {
                Type = OperationType.Update,
                EntityName = entity.LogicalName,
                RecordId = entity.Id
            };

            foreach (var attr in entity.Attributes)
            {
                op.Fields[attr.Key] = FormatValue(attr.Value);
            }

            InterceptedOperations.Add(op);

            // Store in dirty cache so subsequent Retrieves see these changes
            var key = $"{entity.LogicalName}|{entity.Id}";
            if (!_pendingUpdates.ContainsKey(key))
                _pendingUpdates[key] = new Dictionary<string, object>();

            foreach (var attr in entity.Attributes)
            {
                _pendingUpdates[key][attr.Key] = attr.Value;
            }

            LogIntercepted("UPDATE", entity.LogicalName, entity.Id, op.Fields.Count);
        }

        public void Delete(string entityName, Guid id)
        {
            var op = new InterceptedOperation
            {
                Type = OperationType.Delete,
                EntityName = entityName,
                RecordId = id
            };

            InterceptedOperations.Add(op);
            _deletedRecords.Add($"{entityName}|{id}");

            LogIntercepted("DELETE", entityName, id, 0);
        }

        public Entity Retrieve(string entityName, Guid id, ColumnSet columnSet)
        {
            RetrieveCount++;
            var key = $"{entityName}|{id}";

            // Check if this was a record we created during this execution
            if (_createdRecords.ContainsKey(key))
            {
                LogPassThrough("RETRIEVE (from created cache)", entityName, id);
                return _createdRecords[key];
            }

            // Fetch from real D365
            Entity result;
            try
            {
                result = _realService.Retrieve(entityName, id, columnSet);
                LogPassThrough("RETRIEVE", entityName, id);
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"  !! RETRIEVE FAILED: {entityName} ({id}): {ex.Message}");
                Console.ResetColor();
                throw;
            }

            // Merge dirty cache values (overlay pending updates)
            if (_pendingUpdates.ContainsKey(key))
            {
                foreach (var kvp in _pendingUpdates[key])
                {
                    result[kvp.Key] = kvp.Value;
                }
                Console.ForegroundColor = ConsoleColor.DarkYellow;
                Console.WriteLine($"  ↳ Merged {_pendingUpdates[key].Count} dirty cache field(s) into result");
                Console.ResetColor();
            }

            return result;
        }

        public EntityCollection RetrieveMultiple(QueryBase query)
        {
            RetrieveMultipleCount++;

            EntityCollection result;
            try
            {
                result = _realService.RetrieveMultiple(query);
                var queryInfo = GetQueryInfo(query);
                Console.ForegroundColor = ConsoleColor.DarkGray;
                Console.WriteLine($"  → RETRIEVEMULTIPLE: {queryInfo} → {result.Entities.Count} record(s)");
                Console.ResetColor();
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"  !! RETRIEVEMULTIPLE FAILED: {ex.Message}");
                Console.ResetColor();
                throw;
            }

            // Merge dirty cache for any matching records
            foreach (var entity in result.Entities)
            {
                var key = $"{entity.LogicalName}|{entity.Id}";
                if (_pendingUpdates.ContainsKey(key))
                {
                    foreach (var kvp in _pendingUpdates[key])
                    {
                        entity[kvp.Key] = kvp.Value;
                    }
                }
            }

            // Filter out deleted records
            var toRemove = result.Entities
                .Where(e => _deletedRecords.Contains($"{e.LogicalName}|{e.Id}"))
                .ToList();
            foreach (var e in toRemove)
                result.Entities.Remove(e);

            return result;
        }

        public OrganizationResponse Execute(OrganizationRequest request)
        {
            var requestName = request.GetType().Name;

            // Route: is this a write operation?
            if (IsWriteRequest(request))
            {
                return InterceptExecute(request);
            }

            // Read-only request: pass through to real D365
            Console.ForegroundColor = ConsoleColor.DarkGray;
            Console.WriteLine($"  → EXECUTE (pass-through): {requestName}");
            Console.ResetColor();

            try
            {
                return _realService.Execute(request);
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"  !! EXECUTE FAILED ({requestName}): {ex.Message}");
                Console.ResetColor();
                throw;
            }
        }

        public void Associate(string entityName, Guid entityId, Relationship relationship, EntityReferenceCollection relatedEntities)
        {
            var op = new InterceptedOperation
            {
                Type = OperationType.Associate,
                EntityName = entityName,
                RecordId = entityId,
                Details =
                {
                    ["Relationship"] = relationship.SchemaName,
                    ["RelatedEntities"] = string.Join(", ", relatedEntities.Select(r => $"{r.LogicalName}:{r.Id}"))
                }
            };
            InterceptedOperations.Add(op);
            LogIntercepted("ASSOCIATE", entityName, entityId, relatedEntities.Count);
        }

        public void Disassociate(string entityName, Guid entityId, Relationship relationship, EntityReferenceCollection relatedEntities)
        {
            var op = new InterceptedOperation
            {
                Type = OperationType.Disassociate,
                EntityName = entityName,
                RecordId = entityId,
                Details =
                {
                    ["Relationship"] = relationship.SchemaName,
                    ["RelatedEntities"] = string.Join(", ", relatedEntities.Select(r => $"{r.LogicalName}:{r.Id}"))
                }
            };
            InterceptedOperations.Add(op);
            LogIntercepted("DISASSOCIATE", entityName, entityId, relatedEntities.Count);
        }

        #region Private Helpers

        private bool IsWriteRequest(OrganizationRequest request)
        {
            var name = request.GetType().Name;
            if (WriteRequestTypes.Contains(name)) return true;

            // Check by well-known base types
            if (request is CreateRequest || request is UpdateRequest || request is DeleteRequest) return true;

            return false;
        }

        private OrganizationResponse InterceptExecute(OrganizationRequest request)
        {
            var requestName = request.GetType().Name;
            var op = new InterceptedOperation
            {
                Type = OperationType.Execute,
                RequestName = requestName
            };

            // Extract details based on request type
            switch (request)
            {
                case SendEmailRequest sendEmail:
                    op.Details["EmailId"] = sendEmail.EmailId;
                    op.Details["IssueSend"] = sendEmail.IssueSend;
                    break;

                case GrantAccessRequest grant:
                    op.EntityName = grant.Target?.LogicalName;
                    op.RecordId = grant.Target?.Id;
                    op.Details["Principal"] = $"{grant.PrincipalAccess?.Principal?.LogicalName}:{grant.PrincipalAccess?.Principal?.Id}";
                    op.Details["AccessMask"] = grant.PrincipalAccess?.AccessMask.ToString();
                    break;

                case RevokeAccessRequest revoke:
                    op.EntityName = revoke.Target?.LogicalName;
                    op.RecordId = revoke.Target?.Id;
                    op.Details["Revokee"] = $"{revoke.Revokee?.LogicalName}:{revoke.Revokee?.Id}";
                    break;

                case AssignRequest assign:
                    op.EntityName = assign.Target?.LogicalName;
                    op.RecordId = assign.Target?.Id;
                    op.Details["Assignee"] = $"{assign.Assignee?.LogicalName}:{assign.Assignee?.Id}";
                    break;

                case SetStateRequest setState:
                    op.EntityName = setState.EntityMoniker?.LogicalName;
                    op.RecordId = setState.EntityMoniker?.Id;
                    op.Details["State"] = setState.State?.Value.ToString();
                    op.Details["Status"] = setState.Status?.Value.ToString();
                    break;

                default:
                    // Capture all parameters generically
                    foreach (var param in request.Parameters)
                    {
                        op.Details[param.Key] = FormatValue(param.Value);
                    }
                    break;
            }

            InterceptedOperations.Add(op);
            LogIntercepted($"EXECUTE ({requestName})", op.EntityName ?? "N/A", op.RecordId ?? Guid.Empty, op.Details.Count);

            // Return an empty response of the expected type
            return CreateEmptyResponse(request);
        }

        private OrganizationResponse CreateEmptyResponse(OrganizationRequest request)
        {
            switch (request)
            {
                case SendEmailRequest _:
                    return new SendEmailResponse();
                case GrantAccessRequest _:
                    return new GrantAccessResponse();
                case RevokeAccessRequest _:
                    return new RevokeAccessResponse();
                case AssignRequest _:
                    return new AssignResponse();
                case SetStateRequest _:
                    return new SetStateResponse();
                default:
                    return new OrganizationResponse();
            }
        }

        private static object FormatValue(object value)
        {
            switch (value)
            {
                case EntityReference er:
                    return $"EntityRef({er.LogicalName}, {er.Id}, {er.Name})";
                case OptionSetValue osv:
                    return $"OptionSet({osv.Value})";
                case Money m:
                    return $"Money({m.Value})";
                case AliasedValue av:
                    return $"Aliased({av.EntityLogicalName}.{av.AttributeLogicalName}={FormatValue(av.Value)})";
                case Entity e:
                    return $"Entity({e.LogicalName}, {e.Id})";
                case EntityCollection ec:
                    return $"EntityCollection({ec.EntityName}, {ec.Entities.Count} records)";
                case null:
                    return "null";
                default:
                    return value;
            }
        }

        private static string GetQueryInfo(QueryBase query)
        {
            if (query is QueryExpression qe)
                return $"QueryExpression({qe.EntityName})";
            if (query is FetchExpression fe)
            {
                // Extract entity name from FetchXML
                var xml = fe.Query ?? "";
                var start = xml.IndexOf("name='", StringComparison.Ordinal);
                if (start >= 0)
                {
                    start += 6;
                    var end = xml.IndexOf("'", start, StringComparison.Ordinal);
                    if (end > start)
                        return $"FetchXML({xml.Substring(start, end - start)})";
                }
                return "FetchXML(?)";
            }
            return query.GetType().Name;
        }

        private static void LogIntercepted(string operation, string entity, Guid id, int detailCount)
        {
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine($"  ⚡ INTERCEPTED {operation}: {entity} ({id:N}) [{detailCount} field(s)]");
            Console.ResetColor();
        }

        private static void LogPassThrough(string operation, string entity, Guid id)
        {
            Console.ForegroundColor = ConsoleColor.DarkGray;
            Console.WriteLine($"  → {operation}: {entity} ({id:N})");
            Console.ResetColor();
        }

        #endregion
    }
}
