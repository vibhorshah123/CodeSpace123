using System;
using System.Collections.Generic;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace PluginSimulator.Execution
{
    public class AuditEntry
    {
        public Guid AuditId { get; set; }
        public DateTime ChangedOn { get; set; }
        public string ChangedBy { get; set; }
        public Entity OldValue { get; set; }
        public Entity NewValue { get; set; }
        public string ChangedFieldsSummary { get; set; }
    }

    /// <summary>
    /// Fetches audit history for a record to reconstruct pre-image old values.
    /// Requires audit to be enabled on the table in D365.
    /// </summary>
    public static class AuditLoader
    {
        public static List<AuditEntry> LoadRecentChanges(
            IOrganizationService service,
            string entityLogicalName,
            Guid recordId,
            int maxEntries = 50)
        {
            var results = new List<AuditEntry>();

            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine("  Loading audit history...");
            Console.ResetColor();

            string pagingCookie = null;
            int pageNumber = 1;
            const int pageSize = 100; // fetch 100 raw entries per page to find enough AttributeAuditDetail entries

            while (results.Count < maxEntries)
            {
                var paging = new PagingInfo { Count = pageSize, PageNumber = pageNumber };
                if (pagingCookie != null)
                    paging.PagingCookie = pagingCookie;

                var request = new RetrieveRecordChangeHistoryRequest
                {
                    Target = new EntityReference(entityLogicalName, recordId),
                    PagingInfo = paging
                };

                var response = (RetrieveRecordChangeHistoryResponse)service.Execute(request);

                foreach (var detail in response.AuditDetailCollection.AuditDetails)
                {
                    if (!(detail is AttributeAuditDetail attrDetail)) continue;
                    if (results.Count >= maxEntries) break;

                    var changedFields = new List<string>();
                    if (attrDetail.NewValue != null)
                        foreach (var attr in attrDetail.NewValue.Attributes)
                            changedFields.Add(attr.Key);

                    results.Add(new AuditEntry
                    {
                        AuditId              = attrDetail.AuditRecord.Id,
                        ChangedOn            = attrDetail.AuditRecord.GetAttributeValue<DateTime>("createdon"),
                        ChangedBy            = attrDetail.AuditRecord.GetAttributeValue<EntityReference>("userid")?.Name ?? "Unknown",
                        OldValue             = attrDetail.OldValue,
                        NewValue             = attrDetail.NewValue,
                        ChangedFieldsSummary = changedFields.Count > 0
                            ? string.Join(", ", changedFields)
                            : "(no fields)"
                    });
                }

                // Stop if no more pages
                if (!response.AuditDetailCollection.MoreRecords) break;
                pagingCookie = response.AuditDetailCollection.PagingCookie;
                pageNumber++;
            }

            Console.WriteLine($"  Found {results.Count} audit entr{(results.Count == 1 ? "y" : "ies")}");
            return results;
        }

        /// <summary>
        /// Formats a D365 attribute value back to a string suitable for the Changed Fields box.
        /// Pairs with ContextBuilder.ParseFieldValue.
        /// </summary>
        public static string FormatAttributeValue(object value)
        {
            if (value == null) return "";
            if (value is OptionSetValue osv) return osv.Value.ToString();
            if (value is EntityReference er) return er.Id.ToString();
            if (value is Money m) return m.Value.ToString("F2");
            if (value is bool b) return b.ToString().ToLower();
            if (value is DateTime dt) return dt.ToUniversalTime().ToString("o");
            if (value is AliasedValue av) return FormatAttributeValue(av.Value);
            return value.ToString();
        }

        /// <summary>
        /// Serialises a NewValue entity into the "field=value,field2=value2" format for the Changed Fields box.
        /// </summary>
        public static string BuildChangedFieldsString(Entity newValue)
        {
            if (newValue == null) return string.Empty;
            var parts = new List<string>();
            foreach (var attr in newValue.Attributes)
            {
                var formatted = FormatAttributeValue(attr.Value);
                if (!string.IsNullOrEmpty(formatted))
                    parts.Add($"{attr.Key}={formatted}");
            }
            return string.Join(", ", parts);
        }
    }
}
