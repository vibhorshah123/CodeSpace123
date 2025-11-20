using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using Azure.Core;
using Azure.Identity;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Logging;
using Azure.Storage.Blobs;
using ClosedXML.Excel;

namespace D365ComparissionTool;

public class ComparisonRequest
{
    public string? Mode { get; set; } // single | list (default single)
    public string SourceEnvUrl { get; set; } = string.Empty;
    public string TargetEnvUrl { get; set; } = string.Empty;
    public string EntityLogicalName { get; set; } = string.Empty;
    public List<string> Entities { get; set; } = new(); // for list mode
    public List<string> FieldsToCompare { get; set; } = new();
    public bool AutoDiscoverSubgrids { get; set; } // if true and SubgridRelationships empty, discover all one-to-many relationships
    // Direct bearer tokens (preferred when provided)
    public string? SourceBearerToken { get; set; }
    public string? TargetBearerToken { get; set; }
    // Output storage settings
    public string? StorageAccountUrl { get; set; }
    public string? OutputContainerName { get; set; }
    public string? OutputFileName { get; set; }
    // New flexible auth options for storage
    public string? StorageConnectionString { get; set; } // full connection string
    public string? StorageContainerSasUrl { get; set; } // SAS URL directly to container (https://acct.blob.core.windows.net/container?sv=...)

    // Legacy token generation
    public string? TokenGenerationEndpoint { get; set; }
    public string? AppClientId { get; set; }
    public string? ResourceTenantId { get; set; }
    public string? MiClientId { get; set; }
    public string? Scope { get; set; }

    // Source token generation (optional legacy)
    public string? SourceTokenGenerationEndpoint { get; set; }
    public string? SourceAppClientId { get; set; }
    public string? SourceResourceTenantId { get; set; }
    public string? SourceMiClientId { get; set; }
    public string? SourceScope { get; set; }

    public List<SubgridRelationRequest> SubgridRelationships { get; set; } = new();
}

public class SubgridRelationRequest
{
    public string ChildEntityLogicalName { get; set; } = string.Empty; // e.g. mash_childentity
    public string ChildParentFieldLogicalName { get; set; } = string.Empty; // lookup field on child pointing to parent (guid)
    public List<string> ChildFields { get; set; } = new(); // fields to compare on child
    public bool OnlyActiveChildren { get; set; } = true; // apply statecode eq 0 filter
    // Internal enrichment flags
    public bool ParentLookupFieldValid { get; set; } = true; // set false if metadata/selectability validation fails
}

public class MismatchDetail
{
    public Guid RecordId { get; set; }
    public string FieldName { get; set; } = string.Empty;
    public object? SourceValue { get; set; }
    public object? TargetValue { get; set; }
}

public class MultiFieldMismatch
{
    public Guid RecordId { get; set; }
    public List<MismatchDetail> FieldMismatches { get; set; } = new();
}

public class DiffReport
{
    public List<Dictionary<string, object?>> OnlyInSource { get; set; } = new();
    public List<Dictionary<string, object?>> OnlyInTarget { get; set; } = new();
    public List<MismatchDetail> Mismatches { get; set; } = new();
    public List<MultiFieldMismatch> MultiFieldMismatches { get; set; } = new();
    public List<Guid> MatchingRecordIds { get; set; } = new();
    public List<string> ComparedFields { get; set; } = new();
    public string Notes { get; set; } = string.Empty;
    public string? ExcelBlobUrl { get; set; }
    public Dictionary<string, ChildDiffReport> ChildDiffs { get; set; } = new();
    public bool SourceAuthFailed { get; set; } // indicates 401 while fetching source
    public bool TargetAuthFailed { get; set; } // indicates 401 while fetching target
    // Added: primary name attribute for entity (used to output RecordName column consistently in Excel sheets)
    public string? PrimaryNameAttribute { get; set; }
}

public class ChildMismatchDetail
{
    public Guid ParentId { get; set; }
    public Guid ChildId { get; set; }
    public string FieldName { get; set; } = string.Empty;
    public object? SourceValue { get; set; }
    public object? TargetValue { get; set; }
}
public class ChildMultiFieldMismatch
{
    public Guid ParentId { get; set; }
    public Guid ChildId { get; set; }
    public List<ChildMismatchDetail> FieldMismatches { get; set; } = new();
}
public class ChildDiffReport
{
    public string ChildEntityLogicalName { get; set; } = string.Empty;
    public string ParentLookupField { get; set; } = string.Empty;
    public List<string> ComparedChildFields { get; set; } = new();
    public Dictionary<Guid, List<Dictionary<string, object?>>> OnlyInSourceByParent { get; set; } = new();
    public Dictionary<Guid, List<Dictionary<string, object?>>> OnlyInTargetByParent { get; set; } = new();
    public List<ChildMismatchDetail> Mismatches { get; set; } = new();
    public List<ChildMultiFieldMismatch> MultiFieldMismatches { get; set; } = new();
    public int TotalSourceChildRecords { get; set; }
    public int TotalTargetChildRecords { get; set; }
}

public class ConfigurationComparisionFunction
{
    private readonly ILogger _logger;
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true,
        WriteIndented = true
    };

    // Added helper methods for field validation
    private static async Task<bool> IsFieldSelectableAsync(string envUrl, string entityLogicalName, string field, string accessToken)
    {
        try
        {
            using var client = new HttpClient();
            client.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", accessToken);
            client.DefaultRequestHeaders.Accept.ParseAdd("application/json");
            // Resolve correct entity set name (plural) from metadata (fallback to naive pluralization)
            var entitySetName = await GetEntitySetNameAsync(envUrl, entityLogicalName, accessToken) ?? (entityLogicalName + "s");
            var testUrl = envUrl.TrimEnd('/') + $"/api/data/v9.2/{entitySetName}?$select={field}&$top=1";
            using var resp = await client.GetAsync(testUrl);
            return resp.IsSuccessStatusCode;
        }
        catch { return false; }
    }

    private static async Task<List<string>> ValidateSelectableFieldsAsync(string envUrl, string entityLogicalName, List<string> fields, string accessToken, Guid correlationId)
    {
        var validated = new List<string>();
        foreach (var f in fields)
        {
            if (await IsFieldSelectableAsync(envUrl, entityLogicalName, f, accessToken))
            {
                validated.Add(f);
            }
        }
        return validated;
    }

    public ConfigurationComparisionFunction(ILogger<ConfigurationComparisionFunction> logger) => _logger = logger;

    [Function("CompareDataverse")]
    public async Task<HttpResponseData> Run([HttpTrigger(AuthorizationLevel.Function, "post")] HttpRequestData req)
    {
        var correlationId = Guid.NewGuid();
        _logger.LogInformation("[CompareDataverse] Start request CorrelationId={CorrelationId}", correlationId);
        try
        {
            string body = await new System.IO.StreamReader(req.Body).ReadToEndAsync();
            if (string.IsNullOrWhiteSpace(body))
            {
                _logger.LogWarning("[{CorrelationId}] Empty body", correlationId);
                var bad = req.CreateResponse(System.Net.HttpStatusCode.BadRequest);
                await bad.WriteStringAsync("Empty request body");
                return bad;
            }

            var request = JsonSerializer.Deserialize<ComparisonRequest>(body, JsonOptions);
            if (request == null || string.IsNullOrWhiteSpace(request.SourceEnvUrl) || string.IsNullOrWhiteSpace(request.TargetEnvUrl))
            {
                _logger.LogWarning("[{CorrelationId}] Invalid payload SourceEnvUrl='{SourceEnvUrl}' TargetEnvUrl='{TargetEnvUrl}'", correlationId, request?.SourceEnvUrl, request?.TargetEnvUrl);
                var bad = req.CreateResponse(System.Net.HttpStatusCode.BadRequest);
                await bad.WriteStringAsync("Invalid request payload");
                return bad;
            }

            // If caller did not explicitly provide subgrid relationships or disable auto discovery, enable it by default
            if (!request.AutoDiscoverSubgrids && (request.SubgridRelationships == null || request.SubgridRelationships.Count == 0))
            {
                request.AutoDiscoverSubgrids = true; // default behavior: discover subgrids and compare their records
            }

            var mode = string.IsNullOrWhiteSpace(request.Mode) ? "single" : request.Mode.Trim().ToLowerInvariant();
            if (mode == "single")
            {
                if (string.IsNullOrWhiteSpace(request.EntityLogicalName))
                {
                    var bad = req.CreateResponse(System.Net.HttpStatusCode.BadRequest);
                    await bad.WriteStringAsync("single mode requires 'EntityLogicalName'");
                    return bad;
                }
            }
            else if (mode == "list")
            {
                if (request.Entities == null || request.Entities.Count == 0)
                {
                    var bad = req.CreateResponse(System.Net.HttpStatusCode.BadRequest);
                    await bad.WriteStringAsync("list mode requires 'Entities' array");
                    return bad;
                }
            }
            else
            {
                var bad = req.CreateResponse(System.Net.HttpStatusCode.BadRequest);
                await bad.WriteStringAsync("Unsupported mode. Use 'single' or 'list'.");
                return bad;
            }

            _logger.LogInformation("[{CorrelationId}] Request Mode={Mode} EntityOrCount={EntityOrCount} Source={Source} Target={Target} OutputContainer={Container} OutputFileName={FileName}", correlationId, mode, mode == "single" ? request.EntityLogicalName : (request.Entities?.Count.ToString() ?? "0"), request.SourceEnvUrl, request.TargetEnvUrl, request.OutputContainerName, request.OutputFileName);

            var credential = new DefaultAzureCredential();

            // Acquire or use provided tokens
            string sourceToken = !string.IsNullOrWhiteSpace(request.SourceBearerToken)
                ? request.SourceBearerToken.Trim()
                : await AcquireSourceTokenAsync(request, credential);
            if (!string.IsNullOrWhiteSpace(request.SourceBearerToken)) _logger.LogDebug("[{CorrelationId}] Using provided source bearer token length={Len}", correlationId, sourceToken.Length);
            else _logger.LogInformation("[{CorrelationId}] Acquired source token", correlationId);

            string targetToken = !string.IsNullOrWhiteSpace(request.TargetBearerToken)
                ? request.TargetBearerToken.Trim()
                : await AcquireTargetTokenAsync(request, credential);
            if (!string.IsNullOrWhiteSpace(request.TargetBearerToken)) _logger.LogDebug("[{CorrelationId}] Using provided target bearer token length={Len}", correlationId, targetToken.Length);
            else _logger.LogInformation("[{CorrelationId}] Acquired target token", correlationId);

            // Acquire or use provided tokens ONCE and reuse per entity
            _logger.LogInformation("[{CorrelationId}] Discovering mash_ fields", correlationId);
            // Branch by mode
            var singleSourceToken = sourceToken; var singleTargetToken = targetToken;
            if (mode == "single")
            {
                var report = await CompareSingleEntityAsync(request, singleSourceToken, singleTargetToken, credential, correlationId);
                string json = JsonSerializer.Serialize(report, JsonOptions);
                var response = req.CreateResponse(System.Net.HttpStatusCode.OK);
                response.Headers.Add("Content-Type", "application/json");
                await response.WriteStringAsync(json);
                return response;
            }
            else
            {
                var results = new List<(string Entity, DiffReport Report)>();
                foreach (var entity in request.Entities)
                {
                    var original = request.EntityLogicalName;
                    request.EntityLogicalName = entity;
                    var report = await CompareSingleEntityAsync(request, singleSourceToken, singleTargetToken, credential, correlationId);
                    results.Add((entity, report));
                    request.EntityLogicalName = original; // restore
                }
                var payload = new
                {
                    Mode = mode,
                    Results = results.Select(r => new { EntityLogicalName = r.Entity, Report = r.Report }).ToList()
                };
                string json = JsonSerializer.Serialize(payload, JsonOptions);
                var response = req.CreateResponse(System.Net.HttpStatusCode.OK);
                response.Headers.Add("Content-Type", "application/json");
                await response.WriteStringAsync(json);
                return response;
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "[CompareDataverse] Unhandled error CorrelationId={CorrelationId}", correlationId);
            var error = req.CreateResponse(System.Net.HttpStatusCode.InternalServerError);
            await error.WriteStringAsync("Error: " + ex.Message);
            return error;
        }
    }

    private async Task<DiffReport> CompareSingleEntityAsync(ComparisonRequest request, string sourceToken, string targetToken, DefaultAzureCredential credential, Guid correlationId)
    {
        _logger.LogInformation("[{CorrelationId}] Request Entity={Entity} Source={Source} Target={Target} OutputContainer={Container} OutputFileName={FileName}", correlationId, request.EntityLogicalName, request.SourceEnvUrl, request.TargetEnvUrl, request.OutputContainerName, request.OutputFileName);

        // Discover mash_ fields
        _logger.LogInformation("[{CorrelationId}] Discovering mash_ fields", correlationId);
        var sourceFieldTask = GetMashFieldsAsync(request.SourceEnvUrl, request.EntityLogicalName, sourceToken);
        var targetFieldTask = GetMashFieldsAsync(request.TargetEnvUrl, request.EntityLogicalName, targetToken);
        await Task.WhenAll(sourceFieldTask, targetFieldTask);
        var sourceFields = sourceFieldTask.Result;
        var targetFields = targetFieldTask.Result;
        _logger.LogInformation("[{CorrelationId}] Field discovery SourceCount={SourceCount} TargetCount={TargetCount}", correlationId, sourceFields.Count, targetFields.Count);
        var rawIntersection = sourceFields.Intersect(targetFields, StringComparer.OrdinalIgnoreCase).Distinct(StringComparer.OrdinalIgnoreCase).ToList();

        // Discover primary name attribute (for consistent RecordName column). Use source environment metadata.
        var primaryNameAttribute = await GetPrimaryNameAttributeAsync(request.SourceEnvUrl, request.EntityLogicalName, sourceToken);
        if (!string.IsNullOrWhiteSpace(primaryNameAttribute))
        {
            _logger.LogInformation("[{CorrelationId}] PrimaryNameAttribute resolved '{PrimaryNameAttribute}'", correlationId, primaryNameAttribute);
        }
        else
        {
            _logger.LogWarning("[{CorrelationId}] PrimaryNameAttribute not found for entity '{Entity}'", correlationId, request.EntityLogicalName);
        }

        // Normalize: keep base fields; capture formatted name variants mapping
        var formattedVariants = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase); // variant -> base
        var selectableFields = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var f in rawIntersection)
        {
            string baseCandidate = null;
            if (f.EndsWith("typename", StringComparison.OrdinalIgnoreCase))
                baseCandidate = f[..^8]; // remove 'typename'
            else if (f.EndsWith("name", StringComparison.OrdinalIgnoreCase))
                baseCandidate = f[..^4]; // remove 'name'

            if (baseCandidate != null && rawIntersection.Contains(baseCandidate))
            {
                // base exists; map variant -> base and do not include variant in selectable list
                formattedVariants[f] = baseCandidate;
                selectableFields.Add(baseCandidate);
            }
            else
            {
                selectableFields.Add(f);
            }
        }

        // Validate selectable fields individually against API to avoid removing valid base fields erroneously
        var validatedFields = await ValidateSelectableFieldsAsync(request.SourceEnvUrl, request.EntityLogicalName, selectableFields.ToList(), sourceToken, correlationId);
        // Some base fields might fail; if they do and variant existed, attempt to keep variant only (will later derive as null)
        foreach (var kvp in formattedVariants.ToList())
        {
            if (!validatedFields.Contains(kvp.Value))
            {
                // base not selectable; treat variant as ordinary field if it was discovered and validate it alone
                if (rawIntersection.Contains(kvp.Key) && await IsFieldSelectableAsync(request.SourceEnvUrl, request.EntityLogicalName, kvp.Key, sourceToken))
                {
                    validatedFields.Add(kvp.Key);
                    formattedVariants.Remove(kvp.Key); // no longer a pure formatted variant
                    _logger.LogWarning("[{CorrelationId}] Base field '{Base}' invalid; using variant '{Variant}' directly", correlationId, kvp.Value, kvp.Key);
                }
            }
        }

        var fieldsToCompare = rawIntersection; // report original intersection
        _logger.LogInformation("[{CorrelationId}] Intersected fields Count={Count} Fields={Fields}", correlationId, fieldsToCompare.Count, string.Join(',', fieldsToCompare));
        _logger.LogInformation("[{CorrelationId}] Selectable validated fields Count={Count} Fields={Fields}", correlationId, validatedFields.Count, string.Join(',', validatedFields));
        if (formattedVariants.Count > 0)
        {
            _logger.LogInformation("[{CorrelationId}] Formatted variants mapped: {Map}", correlationId, string.Join(';', formattedVariants.Select(m => m.Key + "->" + m.Value)));
        }

        var report = new DiffReport { ComparedFields = fieldsToCompare, PrimaryNameAttribute = primaryNameAttribute };
        if (fieldsToCompare.Count == 0)
        {
            report.Notes = "No mash_ prefixed fields found in both environments for the specified entity.";
            _logger.LogWarning("[{CorrelationId}] No common mash_ fields", correlationId);
        }

        // Fetch records using validated selectable fields (exclude pure formatted variants)
        _logger.LogInformation("[{CorrelationId}] Fetching records", correlationId);
        List<Dictionary<string, object?>> sourceRecords = new();
        List<Dictionary<string, object?>> targetRecords = new();
        try
        {
            sourceRecords = await FetchAllRecordsAsync(_logger, request.SourceEnvUrl, request.EntityLogicalName, validatedFields, sourceToken, onlyActive: true, correlationId, primaryNameAttribute);
        }
        catch (UnauthorizedAccessException uex)
        {
            report.SourceAuthFailed = true;
            report.Notes += (string.IsNullOrEmpty(report.Notes) ? string.Empty : " ") + "Source auth failed: " + uex.Message;
            _logger.LogError(uex, "[{CorrelationId}] Source fetch unauthorized", correlationId);
        }
        try
        {
            targetRecords = await FetchAllRecordsAsync(_logger, request.TargetEnvUrl, request.EntityLogicalName, validatedFields, targetToken, onlyActive: true, correlationId, primaryNameAttribute);
        }
        catch (UnauthorizedAccessException uex)
        {
            report.TargetAuthFailed = true;
            report.Notes += (string.IsNullOrEmpty(report.Notes) ? string.Empty : " ") + "Target auth failed: " + uex.Message;
            _logger.LogError(uex, "[{CorrelationId}] Target fetch unauthorized", correlationId);
        }
        _logger.LogInformation("[{CorrelationId}] Fetched SourceRecords={SourceRecords} TargetRecords={TargetRecords}", correlationId, sourceRecords.Count, targetRecords.Count);

        // Inject formatted variant values (derived from annotations) into record dictionaries
        void InjectFormattedValues(List<Dictionary<string, object?>> records)
        {
            foreach (var rec in records)
            {
                foreach (var variant in formattedVariants)
                {
                    // variant.Key = formatted field logical name (e.g. mash_servicetypename), variant.Value = base field (mash_servicetype)
                    var annotationKey = variant.Value + "@OData.Community.Display.V1.FormattedValue";
                    if (rec.TryGetValue(annotationKey, out var fmt))
                    {
                        rec[variant.Key] = fmt;
                    }
                    else if (!rec.ContainsKey(variant.Key))
                    {
                        rec[variant.Key] = null; // ensure presence with null if missing
                    }
                }
            }
        }
        InjectFormattedValues(sourceRecords);
        InjectFormattedValues(targetRecords);
        _logger.LogInformation("[{CorrelationId}] Fetched SourceRecords={SourceRecords} TargetRecords={TargetRecords}", correlationId, sourceRecords.Count, targetRecords.Count);

        string pkField = request.EntityLogicalName + "id";
        var sourceMap = BuildRecordMap(sourceRecords, pkField);
        var targetMap = BuildRecordMap(targetRecords, pkField);
        _logger.LogDebug("[{CorrelationId}] Built maps SourceMap={SourceMapCount} TargetMap={TargetMapCount}", correlationId, sourceMap.Count, targetMap.Count);

        foreach (var kvp in sourceMap)
        {
            var id = kvp.Key;
            var sourceRecord = kvp.Value;
            if (!targetMap.TryGetValue(id, out var targetRecord))
            {
                report.OnlyInSource.Add(sourceRecord);
                continue;
            }

            int mismatchCount = 0;
            foreach (var field in fieldsToCompare)
            {
                sourceRecord.TryGetValue(field, out var sVal);
                targetRecord.TryGetValue(field, out var tVal);
                // Treat missing field as null implicitly (TryGetValue sets out to null). Empty string -> null normalization done in AreEqual.
                if (!AreEqual(sVal, tVal))
                {
                    mismatchCount++;
                    report.Mismatches.Add(new MismatchDetail { RecordId = id, FieldName = field, SourceValue = sVal, TargetValue = tVal });
                }
            }
            if (mismatchCount == 0) report.MatchingRecordIds.Add(id);
        }

        foreach (var kvp in targetMap)
        {
            if (!sourceMap.ContainsKey(kvp.Key)) report.OnlyInTarget.Add(kvp.Value);
        }

        foreach (var group in report.Mismatches.GroupBy(m => m.RecordId))
        {
            if (group.Count() > 1)
            {
                report.MultiFieldMismatches.Add(new MultiFieldMismatch { RecordId = group.Key, FieldMismatches = group.ToList() });
            }
        }

        _logger.LogInformation("[{CorrelationId}] Diff summary Matching={Matching} MismatchDetails={MismatchDetails} MultiFieldMismatches={MultiFieldMismatchCount} OnlyInSource={OnlySource} OnlyInTarget={OnlyTarget}", correlationId, report.MatchingRecordIds.Count, report.Mismatches.Count, report.MultiFieldMismatches.Count, report.OnlyInSource.Count, report.OnlyInTarget.Count);

        // Subgrid comparisons
        if ((request.SubgridRelationships == null || request.SubgridRelationships.Count == 0) && request.AutoDiscoverSubgrids)
        {
            await AutoDiscoverSubgridRelationshipsAsync(request, sourceToken, correlationId);
        }
        if (request.SubgridRelationships != null && request.SubgridRelationships.Count > 0)
        {
            // Enrich subgrid relations with metadata (auto-populate childFields if empty, validate parent lookup field)
            await EnrichSubgridRelationshipsAsync(request, sourceToken, targetToken, correlationId);
            _logger.LogInformation("[{CorrelationId}] Starting subgrid comparisons Count={Count}", correlationId, request.SubgridRelationships.Count);
            foreach (var rel in request.SubgridRelationships)
            {
                if (string.IsNullOrWhiteSpace(rel.ChildEntityLogicalName) || string.IsNullOrWhiteSpace(rel.ChildParentFieldLogicalName))
                {
                    _logger.LogWarning("[{CorrelationId}] Skipping subgrid relation due to missing child entity or parent field", correlationId);
                    continue;
                }
                if (!rel.ParentLookupFieldValid)
                {
                    _logger.LogDebug("[{CorrelationId}] Skipping subgrid relation ChildEntity={ChildEntity} ParentField={ParentField} due to invalid parent lookup field or empty child field set", correlationId, rel.ChildEntityLogicalName, rel.ChildParentFieldLogicalName);
                    continue;
                }
                if (rel.ChildFields == null || rel.ChildFields.Count == 0)
                {
                    _logger.LogDebug("[{CorrelationId}] Skipping subgrid relation ChildEntity={ChildEntity} - no child fields to compare", correlationId, rel.ChildEntityLogicalName);
                    continue;
                }
                var childReport = new ChildDiffReport
                {
                    ChildEntityLogicalName = rel.ChildEntityLogicalName,
                    ParentLookupField = rel.ChildParentFieldLogicalName,
                    ComparedChildFields = rel.ChildFields.Distinct(StringComparer.OrdinalIgnoreCase).ToList()
                };
                try
                {
                    _logger.LogInformation("[{CorrelationId}] Fetching child records ChildEntity={ChildEntity} ParentField={ParentField}", correlationId, rel.ChildEntityLogicalName, rel.ChildParentFieldLogicalName);
                    var sourceChildren = await FetchChildRecordsAsync(_logger, request.SourceEnvUrl, rel.ChildEntityLogicalName, rel.ChildParentFieldLogicalName, rel.ChildFields, sourceToken, rel.OnlyActiveChildren, correlationId);
                    var targetChildren = await FetchChildRecordsAsync(_logger, request.TargetEnvUrl, rel.ChildEntityLogicalName, rel.ChildParentFieldLogicalName, rel.ChildFields, targetToken, rel.OnlyActiveChildren, correlationId);
                    childReport.TotalSourceChildRecords = sourceChildren.Count;
                    childReport.TotalTargetChildRecords = targetChildren.Count;

                    string childPk = rel.ChildEntityLogicalName + "id";
                    // Group by parent
                    var sourceByParent = sourceChildren.Where(c => c.TryGetValue(rel.ChildParentFieldLogicalName, out _)).GroupBy(c => Guid.TryParse(c[rel.ChildParentFieldLogicalName]?.ToString(), out var g) ? g : Guid.Empty).Where(g => g.Key != Guid.Empty).ToDictionary(g => g.Key, g => g.ToList());
                    var targetByParent = targetChildren.Where(c => c.TryGetValue(rel.ChildParentFieldLogicalName, out _)).GroupBy(c => Guid.TryParse(c[rel.ChildParentFieldLogicalName]?.ToString(), out var g) ? g : Guid.Empty).Where(g => g.Key != Guid.Empty).ToDictionary(g => g.Key, g => g.ToList());

                    var allParentIds = new HashSet<Guid>(sourceByParent.Keys.Concat(targetByParent.Keys));
                    foreach (var parentId in allParentIds)
                    {
                        sourceByParent.TryGetValue(parentId, out var srcList);
                        targetByParent.TryGetValue(parentId, out var tgtList);
                        var srcMap = (srcList ?? new List<Dictionary<string, object?>>()).Where(r => r.TryGetValue(childPk, out _)).ToDictionary(r => Guid.TryParse(r[childPk]?.ToString(), out var g) ? g : Guid.Empty, r => r);
                        var tgtMap = (tgtList ?? new List<Dictionary<string, object?>>()).Where(r => r.TryGetValue(childPk, out _)).ToDictionary(r => Guid.TryParse(r[childPk]?.ToString(), out var g) ? g : Guid.Empty, r => r);
                        srcMap.Remove(Guid.Empty); tgtMap.Remove(Guid.Empty);

                        // OnlyInSource / OnlyInTarget children
                        var onlySourceChildren = new List<Dictionary<string, object?>>();
                        foreach (var kv in srcMap)
                            if (!tgtMap.ContainsKey(kv.Key)) onlySourceChildren.Add(kv.Value);
                        var onlyTargetChildren = new List<Dictionary<string, object?>>();
                        foreach (var kv in tgtMap)
                            if (!srcMap.ContainsKey(kv.Key)) onlyTargetChildren.Add(kv.Value);
                        if (onlySourceChildren.Count > 0) childReport.OnlyInSourceByParent[parentId] = onlySourceChildren;
                        if (onlyTargetChildren.Count > 0) childReport.OnlyInTargetByParent[parentId] = onlyTargetChildren;

                        // Mismatches
                        foreach (var kv in srcMap)
                        {
                            if (!tgtMap.TryGetValue(kv.Key, out var tgtChild)) continue;
                            int childMismatchCounter = 0;
                            foreach (var f in childReport.ComparedChildFields)
                            {
                                kv.Value.TryGetValue(f, out var sv);
                                tgtChild.TryGetValue(f, out var tv);
                                if (!AreEqual(sv, tv))
                                {
                                    childMismatchCounter++;
                                    var md = new ChildMismatchDetail { ParentId = parentId, ChildId = kv.Key, FieldName = f, SourceValue = sv, TargetValue = tv };
                                    childReport.Mismatches.Add(md);
                                }
                            }
                            if (childMismatchCounter > 1)
                            {
                                childReport.MultiFieldMismatches.Add(new ChildMultiFieldMismatch
                                {
                                    ParentId = parentId,
                                    ChildId = kv.Key,
                                    FieldMismatches = childReport.Mismatches.Where(m => m.ParentId == parentId && m.ChildId == kv.Key).ToList()
                                });
                            }
                        }
                    }
                    report.ChildDiffs[rel.ChildEntityLogicalName] = childReport;
                    _logger.LogInformation("[{CorrelationId}] Child diff completed ChildEntity={ChildEntity} SourceChildren={SrcCount} TargetChildren={TgtCount} Mismatches={MismatchCount}", correlationId, rel.ChildEntityLogicalName, childReport.TotalSourceChildRecords, childReport.TotalTargetChildRecords, childReport.Mismatches.Count);
                }
                catch (Exception childEx)
                {
                    _logger.LogError(childEx, "[{CorrelationId}] Failed child relation comparison ChildEntity={ChildEntity}", correlationId, rel.ChildEntityLogicalName);
                    report.Notes += (string.IsNullOrEmpty(report.Notes) ? string.Empty : " ") + $"Child comparison failed for {rel.ChildEntityLogicalName}: {childEx.Message}";
                }
            }
        }

        // Excel upload (optional)
        string storageUrl = request.StorageAccountUrl ?? Environment.GetEnvironmentVariable("COMPARISON_STORAGE_URL") ?? string.Empty;
        string containerName = request.OutputContainerName ?? Environment.GetEnvironmentVariable("COMPARISON_OUTPUT_CONTAINER") ?? "comparissiontooloutput";
        if (!string.IsNullOrWhiteSpace(request.StorageContainerSasUrl))
        {
            _logger.LogInformation("[{CorrelationId}] Using SAS URL for upload", correlationId);
            try
            {
                string blobName = GenerateBlobName(request.EntityLogicalName);
                var excelBytes = BuildExcel(report, request.EntityLogicalName);
                _logger.LogDebug("[{CorrelationId}] Excel size bytes={Size}", correlationId, excelBytes.Length);
                report.ExcelBlobUrl = await UploadExcelWithFlexibleAuthAsync(excelBytes, blobName, containerSasUrl: request.StorageContainerSasUrl);
                _logger.LogInformation("[{CorrelationId}] Excel uploaded via SAS BlobUrl={BlobUrl}", correlationId, report.ExcelBlobUrl);
            }
            catch (Exception upEx)
            {
                _logger.LogError(upEx, "[{CorrelationId}] Excel upload failed (SAS)", correlationId);
                report.Notes += (string.IsNullOrEmpty(report.Notes) ? string.Empty : " ") + "Excel upload failed (SAS): " + upEx.Message;
            }
        }
        else if (!string.IsNullOrWhiteSpace(request.StorageConnectionString))
        {
            _logger.LogInformation("[{CorrelationId}] Using connection string for upload", correlationId);
            try
            {
                string blobName = GenerateBlobName(request.EntityLogicalName);
                var excelBytes = BuildExcel(report, request.EntityLogicalName);
                _logger.LogDebug("[{CorrelationId}] Excel size bytes={Size}", correlationId, excelBytes.Length);
                report.ExcelBlobUrl = await UploadExcelWithFlexibleAuthAsync(excelBytes, blobName, connectionString: request.StorageConnectionString, containerName: containerName);
                _logger.LogInformation("[{CorrelationId}] Excel uploaded via connection string BlobUrl={BlobUrl}", correlationId, report.ExcelBlobUrl);
            }
            catch (Exception upEx)
            {
                _logger.LogError(upEx, "[{CorrelationId}] Excel upload failed (ConnStr)", correlationId);
                report.Notes += (string.IsNullOrEmpty(report.Notes) ? string.Empty : " ") + "Excel upload failed (ConnStr): " + upEx.Message;
            }
        }
        else if (!string.IsNullOrWhiteSpace(storageUrl))
        {
            _logger.LogInformation("[{CorrelationId}] Using managed identity for upload AccountUrl={AccountUrl} Container={Container}", correlationId, storageUrl, containerName);
            try
            {
                string blobName = GenerateBlobName(request.EntityLogicalName);
                var excelBytes = BuildExcel(report, request.EntityLogicalName);
                _logger.LogDebug("[{CorrelationId}] Excel size bytes={Size}", correlationId, excelBytes.Length);
                report.ExcelBlobUrl = await UploadExcelWithFlexibleAuthAsync(excelBytes, blobName, accountUrl: storageUrl, containerName: containerName, credential: credential);
                _logger.LogInformation("[{CorrelationId}] Excel uploaded via identity BlobUrl={BlobUrl}", correlationId, report.ExcelBlobUrl);
            }
            catch (Exception upEx)
            {
                _logger.LogError(upEx, "[{CorrelationId}] Excel upload failed (Identity)", correlationId);
                report.Notes += (string.IsNullOrEmpty(report.Notes) ? string.Empty : " ") + "Excel upload failed (Identity): " + upEx.Message;
            }
        }
        else
        {
            _logger.LogWarning("[{CorrelationId}] No storage configuration provided; Excel skipped", correlationId);
            report.Notes += (string.IsNullOrEmpty(report.Notes) ? string.Empty : " ") + "No storage configuration provided; Excel not generated.";
        }

        _logger.LogInformation("[{CorrelationId}] Comparison completed Entity={Entity} ExcelUrl={ExcelUrl}", correlationId, request.EntityLogicalName, report.ExcelBlobUrl ?? "(none)");
        return report;
    }

    // Discover one-to-many relationships for the parent entity and populate SubgridRelationships if empty
    private async Task AutoDiscoverSubgridRelationshipsAsync(ComparisonRequest request, string sourceToken, Guid correlationId)
    {
        if (string.IsNullOrWhiteSpace(request.EntityLogicalName)) return;
        try
        {
            _logger.LogInformation("[{CorrelationId}] Auto-discovering subgrid relationships for Entity={Entity}", correlationId, request.EntityLogicalName);
            using var client = new HttpClient();
            client.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", sourceToken);
            client.DefaultRequestHeaders.Accept.ParseAdd("application/json");
            string url = request.SourceEnvUrl.TrimEnd('/') + $"/api/data/v9.2/EntityDefinitions(LogicalName='{request.EntityLogicalName}')/OneToManyRelationships?$select=ReferencingEntity,ReferencingAttribute,SchemaName";
            using var resp = await client.GetAsync(url);
            if (!resp.IsSuccessStatusCode)
            {
                var body = await resp.Content.ReadAsStringAsync();
                _logger.LogWarning("[{CorrelationId}] Failed to query relationships status={Status} snippet={Snippet}", correlationId, resp.StatusCode, body.Length > 200 ? body[..200] : body);
                return;
            }
            using var stream = await resp.Content.ReadAsStreamAsync();
            using var doc = await JsonDocument.ParseAsync(stream);
            var list = new List<SubgridRelationRequest>();
            if (doc.RootElement.TryGetProperty("value", out var arr))
            {
                foreach (var rel in arr.EnumerateArray())
                {
                    string referencingEntity = rel.TryGetProperty("ReferencingEntity", out var re) ? re.GetString() ?? string.Empty : string.Empty;
                    string referencingAttribute = rel.TryGetProperty("ReferencingAttribute", out var ra) ? ra.GetString() ?? string.Empty : string.Empty;
                    if (string.IsNullOrWhiteSpace(referencingEntity) || string.IsNullOrWhiteSpace(referencingAttribute)) continue;
                    if (referencingEntity.StartsWith("msdyn_", StringComparison.OrdinalIgnoreCase)) continue;
                    var relReq = new SubgridRelationRequest
                    {
                        ChildEntityLogicalName = referencingEntity,
                        ChildParentFieldLogicalName = referencingAttribute,
                        OnlyActiveChildren = true
                    };
                    list.Add(relReq);
                }
            }
            request.SubgridRelationships = list.GroupBy(l => (l.ChildEntityLogicalName.ToLowerInvariant(), l.ChildParentFieldLogicalName.ToLowerInvariant()))
                .Select(g => g.First()).ToList();
            _logger.LogInformation("[{CorrelationId}] Auto-discovered SubgridRelationships Count={Count}", correlationId, request.SubgridRelationships.Count);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "[{CorrelationId}] Auto-discovery of subgrids failed", correlationId);
        }
    }

    // Populate childFields (if empty) with mash_ intersection and validate parent lookup field existence via metadata
    private async Task EnrichSubgridRelationshipsAsync(ComparisonRequest request, string sourceToken, string targetToken, Guid correlationId)
    {
        foreach (var rel in request.SubgridRelationships)
        {
            if (string.IsNullOrWhiteSpace(rel.ChildEntityLogicalName)) continue;
            // Auto populate child fields if not provided or empty
            if (rel.ChildFields == null || rel.ChildFields.Count == 0)
            {
                try
                {
                    var sourceChildFieldsTask = GetMashFieldsAsync(request.SourceEnvUrl, rel.ChildEntityLogicalName, sourceToken);
                    var targetChildFieldsTask = GetMashFieldsAsync(request.TargetEnvUrl, rel.ChildEntityLogicalName, targetToken);
                    await Task.WhenAll(sourceChildFieldsTask, targetChildFieldsTask);
                    var sourceChildFields = sourceChildFieldsTask.Result;
                    var targetChildFields = targetChildFieldsTask.Result;
                    var intersection = sourceChildFields.Intersect(targetChildFields, StringComparer.OrdinalIgnoreCase)
                        .Distinct(StringComparer.OrdinalIgnoreCase).ToList();
                    // Normalize: remove formatted display variants (ending in 'name' or 'typename') when base exists (e.g. mash_areaname -> mash_area)
                    var baseSet = new HashSet<string>(intersection, StringComparer.OrdinalIgnoreCase);
                    var normalized = new List<string>();
                    foreach (var f in intersection)
                    {
                        string? baseCandidate = null;
                        if (f.EndsWith("typename", StringComparison.OrdinalIgnoreCase)) baseCandidate = f[..^8];
                        else if (f.EndsWith("name", StringComparison.OrdinalIgnoreCase)) baseCandidate = f[..^4];
                        if (baseCandidate != null && baseSet.Contains(baseCandidate))
                        {
                            // Skip the variant; base field will allow retrieval + formatted value annotation
                            continue;
                        }
                        normalized.Add(f);
                    }
                    // Validate selectability of each normalized child field against API (avoid requesting non-existent properties like mash_areaname)
                    var selectable = new List<string>();
                    foreach (var f in normalized)
                    {
                        if (await IsFieldSelectableAsync(request.SourceEnvUrl, rel.ChildEntityLogicalName, f, sourceToken)) selectable.Add(f);
                        else _logger.LogDebug("[{CorrelationId}] Child field '{Field}' not selectable and will be excluded ChildEntity={ChildEntity}", correlationId, f, rel.ChildEntityLogicalName);
                    }
                    rel.ChildFields = selectable;
                    _logger.LogInformation("[{CorrelationId}] Auto-populated normalized selectable childFields for ChildEntity={ChildEntity} Original={OrigCount} Normalized={NormCount} Selectable={SelCount}", correlationId, rel.ChildEntityLogicalName, intersection.Count, normalized.Count, rel.ChildFields.Count);
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "[{CorrelationId}] Failed to auto-populate child fields ChildEntity={ChildEntity}", correlationId, rel.ChildEntityLogicalName);
                }
            }

            // Validate parent lookup field existence in source metadata (basic check)
            if (!string.IsNullOrWhiteSpace(rel.ChildParentFieldLogicalName))
            {
                bool exists = await FieldExistsInMetadataAsync(request.SourceEnvUrl, rel.ChildEntityLogicalName, rel.ChildParentFieldLogicalName, sourceToken);
                if (!exists)
                {
                    rel.ParentLookupFieldValid = false;
                    _logger.LogWarning("[{CorrelationId}] Parent lookup field '{Field}' not found in metadata for child entity '{ChildEntity}' - relation will be skipped", correlationId, rel.ChildParentFieldLogicalName, rel.ChildEntityLogicalName);
                }
                else
                {
                    // Field exists in metadata; test selectability (sometimes relationship attribute not directly selectable)
                    bool selectable = await IsFieldSelectableAsync(request.SourceEnvUrl, rel.ChildEntityLogicalName, rel.ChildParentFieldLogicalName, sourceToken);
                    if (!selectable)
                    {
                        rel.ParentLookupFieldValid = false;
                        _logger.LogWarning("[{CorrelationId}] Parent lookup field '{Field}' for child entity '{ChildEntity}' not selectable - relation will be skipped", correlationId, rel.ChildParentFieldLogicalName, rel.ChildEntityLogicalName);
                    }
                    else
                    {
                        rel.ParentLookupFieldValid = true;
                    }
                }
            }
            else
            {
                rel.ParentLookupFieldValid = false;
                _logger.LogWarning("[{CorrelationId}] Missing parent lookup field for child entity '{ChildEntity}' - relation will be skipped", correlationId, rel.ChildEntityLogicalName);
            }

            // Mark relation invalid if no child fields discovered
            if (rel.ChildFields == null || rel.ChildFields.Count == 0)
            {
                rel.ParentLookupFieldValid = false;
                _logger.LogDebug("[{CorrelationId}] Child entity '{ChildEntity}' has zero selectable mash_ fields after normalization - relation marked invalid", correlationId, rel.ChildEntityLogicalName);
            }
        }
    }

    private static async Task<bool> FieldExistsInMetadataAsync(string envUrl, string entityLogicalName, string fieldLogicalName, string accessToken)
    {
        try
        {
            using var client = new HttpClient();
            client.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", accessToken);
            client.DefaultRequestHeaders.Accept.ParseAdd("application/json");
            // Filter directly for the field name to reduce payload
            string url = envUrl.TrimEnd('/') + $"/api/data/v9.2/EntityDefinitions(LogicalName='{entityLogicalName}')/Attributes?$select=LogicalName&$filter=LogicalName eq '{fieldLogicalName.Replace("'", "''")}'";
            using var resp = await client.GetAsync(url);
            if (!resp.IsSuccessStatusCode) return false;
            using var stream = await resp.Content.ReadAsStreamAsync();
            using var doc = await JsonDocument.ParseAsync(stream);
            if (doc.RootElement.TryGetProperty("value", out var arr))
            {
                foreach (var item in arr.EnumerateArray())
                {
                    if (item.TryGetProperty("LogicalName", out var ln) && string.Equals(ln.GetString(), fieldLogicalName, StringComparison.OrdinalIgnoreCase))
                        return true;
                }
            }
            return false;
        }
        catch { return false; }
    }

    private static async Task<string> AcquireSourceTokenAsync(ComparisonRequest request, DefaultAzureCredential fallbackCredential)
    {
        if (!string.IsNullOrWhiteSpace(request.SourceTokenGenerationEndpoint) && !string.IsNullOrWhiteSpace(request.SourceAppClientId) && !string.IsNullOrWhiteSpace(request.SourceResourceTenantId) && !string.IsNullOrWhiteSpace(request.SourceMiClientId))
        {
            string scope = string.IsNullOrWhiteSpace(request.SourceScope) ? request.SourceEnvUrl.TrimEnd('/') + "/.default" : request.SourceScope;
            var payload = new { AppClientId = request.SourceAppClientId, ResourceTenantId = request.SourceResourceTenantId, MiClientId = request.SourceMiClientId, Scope = scope };
            using var client = new HttpClient();
            using var resp = await client.PostAsync(request.SourceTokenGenerationEndpoint, new StringContent(JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json"));
            resp.EnsureSuccessStatusCode();
            return ParseTokenResponse(await resp.Content.ReadAsStringAsync());
        }
        var token = await fallbackCredential.GetTokenAsync(new TokenRequestContext(new[] { request.SourceEnvUrl.TrimEnd('/') + "/.default" }));
        return token.Token;
    }

    private static async Task<string> AcquireTargetTokenAsync(ComparisonRequest request, DefaultAzureCredential fallbackCredential)
    {
        if (!string.IsNullOrWhiteSpace(request.TokenGenerationEndpoint) && !string.IsNullOrWhiteSpace(request.AppClientId) && !string.IsNullOrWhiteSpace(request.ResourceTenantId) && !string.IsNullOrWhiteSpace(request.MiClientId))
        {
            string scope = string.IsNullOrWhiteSpace(request.Scope) ? request.TargetEnvUrl.TrimEnd('/') + "/.default" : request.Scope;
            var payload = new { request.AppClientId, request.ResourceTenantId, request.MiClientId, Scope = scope };
            using var client = new HttpClient();
            using var resp = await client.PostAsync(request.TokenGenerationEndpoint, new StringContent(JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json"));
            resp.EnsureSuccessStatusCode();
            return ParseTokenResponse(await resp.Content.ReadAsStringAsync());
        }
        var token = await fallbackCredential.GetTokenAsync(new TokenRequestContext(new[] { request.TargetEnvUrl.TrimEnd('/') + "/.default" }));
        return token.Token;
    }

    private static string ParseTokenResponse(string raw)
    {
        try
        {
            using var doc = JsonDocument.Parse(raw);
            if (doc.RootElement.TryGetProperty("access_token", out var at)) return at.GetString() ?? string.Empty;
            if (doc.RootElement.TryGetProperty("token", out var t)) return t.GetString() ?? string.Empty;
        }
        catch { }
        return raw.Trim().Trim('"');
    }

    private static async Task<List<string>> GetMashFieldsAsync(string envUrl, string entityLogicalName, string accessToken)
    {
        var client = new HttpClient();
        client.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", accessToken);
        client.DefaultRequestHeaders.Accept.ParseAdd("application/json");
        string url = envUrl.TrimEnd('/') + $"/api/data/v9.2/EntityDefinitions(LogicalName='{entityLogicalName}')/Attributes?$select=LogicalName";
        var result = new List<string>();
        try
        {
            while (!string.IsNullOrEmpty(url))
            {
                using var req = new HttpRequestMessage(HttpMethod.Get, url);
                using var resp = await client.SendAsync(req);
                if (!resp.IsSuccessStatusCode) return result;
                using var stream = await resp.Content.ReadAsStreamAsync();
                using var doc = await JsonDocument.ParseAsync(stream);
                if (doc.RootElement.TryGetProperty("value", out var arr))
                {
                    foreach (var item in arr.EnumerateArray())
                    {
                        if (item.TryGetProperty("LogicalName", out var ln))
                        {
                            var name = ln.GetString();
                            if (!string.IsNullOrWhiteSpace(name) && name.StartsWith("mash_", StringComparison.OrdinalIgnoreCase)) result.Add(name);
                        }
                    }
                }
                url = doc.RootElement.TryGetProperty("@odata.nextLink", out var next) ? next.GetString() : null;
            }
        }
        catch { }
        return result.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
    }

    private static async Task<List<Dictionary<string, object?>>> FetchAllRecordsAsync(ILogger logger, string envUrl, string entityLogicalName, IEnumerable<string> fields, string accessToken, bool onlyActive, Guid correlationId, string? primaryNameAttribute)
    {
        var client = new HttpClient();
        client.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", accessToken);
        client.DefaultRequestHeaders.Accept.ParseAdd("application/json");
        client.DefaultRequestHeaders.Add("Prefer", "odata.include-annotations=\"OData.Community.Display.V1.FormattedValue\"");

        string pkField = entityLogicalName + "id";
        var selectFields = new List<string> { pkField };
        if (!string.IsNullOrWhiteSpace(primaryNameAttribute)) selectFields.Add(primaryNameAttribute);
        selectFields.AddRange(fields);
        selectFields = selectFields.Where(s => !string.IsNullOrWhiteSpace(s)).Distinct(StringComparer.OrdinalIgnoreCase).Where(s => s.Length < 129).ToList();
        string entitySetName = entityLogicalName + "s"; // TODO: replace with real EntitySetName lookup if needed
        string baseUrl = envUrl.TrimEnd('/') + $"/api/data/v9.2/{entitySetName}";
        string filter = onlyActive ? "&$filter=statecode eq 0" : string.Empty;

        var all = new List<Dictionary<string, object?>>();
        int attempt = 0;
        while (true)
        {
            attempt++;
            bool restartRequested = false;
            string selectClause = string.Join(',', selectFields);
            string url = baseUrl + "?$select=" + selectClause + filter;
            logger.LogDebug("[{CorrelationId}] Fetch attempt {Attempt} URL length={UrlLength} Fields={FieldCount} Select={Select}", correlationId, attempt, url.Length, selectFields.Count, selectClause);
            try
            {
                while (!string.IsNullOrEmpty(url))
                {
                    using var request = new HttpRequestMessage(HttpMethod.Get, url);
                    using var response = await client.SendAsync(request);
                    if (!response.IsSuccessStatusCode)
                    {
                        if ((int)response.StatusCode == 401)
                        {
                            var body = await response.Content.ReadAsStringAsync();
                            logger.LogError("[{CorrelationId}] 401 Unauthorized fetching {Entity} attempt={Attempt} snippet={Snippet}", correlationId, entityLogicalName, attempt, body.Length > 200 ? body[..200] : body);
                            throw new UnauthorizedAccessException($"Unauthorized for {entityLogicalName}. Ensure token audience matches environment URL.");
                        }
                        if ((int)response.StatusCode == 400)
                        {
                            var content = await response.Content.ReadAsStringAsync();
                            logger.LogWarning("[{CorrelationId}] 400 response attempt {Attempt} snippet={Snippet}", correlationId, attempt, content.Length > 300 ? content[..300] : content);
                            var invalidProps = ExtractInvalidProperties(content);
                            if (invalidProps.Count > 0)
                            {
                                int removed = 0;
                                foreach (var ip in invalidProps)
                                {
                                    if (selectFields.Remove(ip)) { removed++; logger.LogWarning("[{CorrelationId}] Removed invalid field '{Field}'", correlationId, ip); }
                                }
                                if (removed > 0)
                                {
                                    restartRequested = true;
                                    break; // restart with reduced field set
                                }
                            }
                            if (invalidProps.Count == 0 && selectFields.Count > 1)
                            {
                                logger.LogWarning("[{CorrelationId}] Could not extract invalid field; falling back to primary key only", correlationId);
                                selectFields = new List<string> { pkField }; restartRequested = true; break;
                            }
                        }
                        else
                        {
                            response.EnsureSuccessStatusCode();
                        }
                    }
                    using var stream = await response.Content.ReadAsStreamAsync();
                    using var doc = await JsonDocument.ParseAsync(stream);
                    if (doc.RootElement.TryGetProperty("value", out var valueArray))
                    {
                        foreach (var item in valueArray.EnumerateArray()) all.Add(ToDictionary(item));
                    }
                    url = doc.RootElement.TryGetProperty("@odata.nextLink", out var nextLink) ? nextLink.GetString() : null;
                }
                if (!restartRequested) break;
            }
            catch (UnauthorizedAccessException) { throw; }
            catch (Exception ex) when (attempt < 5)
            {
                logger.LogWarning(ex, "[{CorrelationId}] Attempt {Attempt} failed; will retry if possible", correlationId, attempt);
            }
            if (!restartRequested || selectFields.Count == 0 || attempt >= 8) break;
        }
        logger.LogInformation("[{CorrelationId}] Fetch completed Records={RecordCount} FinalFieldCount={FieldCount} FinalFields={Fields}", correlationId, all.Count, selectFields.Count, string.Join(',', selectFields));
        return all;

        static List<string> ExtractInvalidProperties(string body)
        {
            var list = new List<string>();
            string marker1 = "Could not find a property named '";
            int idx = body.IndexOf(marker1, StringComparison.OrdinalIgnoreCase);
            while (idx >= 0)
            {
                int start = idx + marker1.Length;
                int end = body.IndexOf("'", start);
                if (end > start)
                {
                    list.Add(body[start..end]);
                    idx = body.IndexOf(marker1, end, StringComparison.OrdinalIgnoreCase);
                }
                else break;
            }
            string marker2 = "Property '";
            idx = body.IndexOf(marker2, StringComparison.OrdinalIgnoreCase);
            while (idx >= 0)
            {
                int start = idx + marker2.Length;
                int end = body.IndexOf("'", start);
                if (end > start)
                {
                    var val = body[start..end];
                    if (!list.Contains(val, StringComparer.OrdinalIgnoreCase)) list.Add(val);
                    idx = body.IndexOf(marker2, end, StringComparison.OrdinalIgnoreCase);
                }
                else break;
            }
            return list.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        }
    }

    private static Dictionary<Guid, Dictionary<string, object?>> BuildRecordMap(List<Dictionary<string, object?>> records, string pkField)
    {
        var dict = new Dictionary<Guid, Dictionary<string, object?>>();
        foreach (var record in records)
        {
            if (record.TryGetValue(pkField, out var idObj) && idObj != null && Guid.TryParse(idObj.ToString(), out var id)) dict[id] = record;
        }
        return dict;
    }

    private static Dictionary<string, object?> ToDictionary(JsonElement element)
    {
        var dict = new Dictionary<string, object?>();
        foreach (var prop in element.EnumerateObject()) dict[prop.Name] = ExtractValue(prop.Value);
        return dict;
    }

    private static object? ExtractValue(JsonElement el) => el.ValueKind switch
    {
        JsonValueKind.String => el.GetString(),
        JsonValueKind.Number => el.TryGetInt64(out var l) ? l : el.TryGetDouble(out var d) ? d : el.GetRawText(),
        JsonValueKind.True => true,
        JsonValueKind.False => false,
        JsonValueKind.Null => null,
        JsonValueKind.Object => JsonSerializer.Deserialize<Dictionary<string, object?>>(el.GetRawText()),
        JsonValueKind.Array => JsonSerializer.Deserialize<List<object?>>(el.GetRawText()),
        _ => el.GetRawText()
    };

    private static bool AreEqual(object? a, object? b)
    {
        // Treat missing or empty string values as null for comparison
        if (a is string sa && string.IsNullOrWhiteSpace(sa)) a = null;
        if (b is string sb && string.IsNullOrWhiteSpace(sb)) b = null;
        if (a == null && b == null) return true;
        if (a == null || b == null) return false;
        if (Guid.TryParse(a.ToString(), out var ga) && Guid.TryParse(b.ToString(), out var gb)) return ga.Equals(gb);
        return string.Equals(a.ToString(), b.ToString(), StringComparison.OrdinalIgnoreCase);
    }

    private static byte[] BuildExcel(DiffReport report, string entityLogicalName)
    {
        const int ExcelMaxCellLength = 32767;
        string Truncate(string? v)
        {
            if (string.IsNullOrEmpty(v)) return v ?? string.Empty;
            if (v.Length <= ExcelMaxCellLength) return v;
            return v.Substring(0, ExcelMaxCellLength - 15) + "...(truncated)";
        }

        using var wb = new XLWorkbook();
        var summary = wb.Worksheets.Add("Summary");
        summary.Cell(1, 1).Value = "Entity"; summary.Cell(1, 2).Value = entityLogicalName;
        summary.Cell(2, 1).Value = "Compared Fields"; summary.Cell(2, 2).Value = Truncate(string.Join(',', report.ComparedFields));
        summary.Cell(3, 1).Value = "Matching Records"; summary.Cell(3, 2).Value = report.MatchingRecordIds.Count;
        summary.Cell(4, 1).Value = "Mismatches"; summary.Cell(4, 2).Value = report.Mismatches.Count; // each field mismatch already listed individually
        summary.Cell(5, 1).Value = "Only In Source"; summary.Cell(5, 2).Value = report.OnlyInSource.Count;
        summary.Cell(6, 1).Value = "Only In Target"; summary.Cell(6, 2).Value = report.OnlyInTarget.Count;
        summary.Cell(7, 1).Value = "Notes"; summary.Cell(7, 2).Value = Truncate(report.Notes);
        summary.Columns().AdjustToContents();

        var matchWs = wb.Worksheets.Add("MatchingRecords");
        matchWs.Cell(1, 1).Value = "RecordId";
        for (int i = 0; i < report.MatchingRecordIds.Count; i++) matchWs.Cell(i + 2, 1).Value = report.MatchingRecordIds[i].ToString();
        matchWs.Columns().AdjustToContents();

        var mismatchWs = wb.Worksheets.Add("Mismatches");
        mismatchWs.Cell(1, 1).Value = "RecordId";
        mismatchWs.Cell(1, 2).Value = "FieldName";
        mismatchWs.Cell(1, 3).Value = "SourceValue";
        mismatchWs.Cell(1, 4).Value = "TargetValue";
        int r = 2;
        foreach (var mm in report.Mismatches)
        {
            mismatchWs.Cell(r, 1).Value = mm.RecordId.ToString();
            mismatchWs.Cell(r, 2).Value = mm.FieldName;
            mismatchWs.Cell(r, 3).Value = Truncate(mm.SourceValue?.ToString());
            mismatchWs.Cell(r, 4).Value = Truncate(mm.TargetValue?.ToString());
            r++;
        }
        mismatchWs.Columns().AdjustToContents();

        // Removed separate MultiFieldMismatches worksheet per request; individual field mismatches already captured above.

        // Write OnlyInSource / OnlyInTarget sheets with just RecordId column as requested
        void WriteRecordIdOnly(string sheetName, List<Dictionary<string, object?>> records)
        {
            var ws = wb.Worksheets.Add(sheetName);
            ws.Cell(1, 1).Value = "RecordId";
            int row = 2;
            string pk = entityLogicalName + "id";
            foreach (var rec in records)
            {
                if (rec.TryGetValue(pk, out var idVal))
                {
                    ws.Cell(row, 1).Value = idVal?.ToString();
                    row++;
                }
            }
            ws.Columns().AdjustToContents();
        }
        WriteRecordIdOnly("OnlyInSource", report.OnlyInSource);
        WriteRecordIdOnly("OnlyInTarget", report.OnlyInTarget);

        // Child diffs (unchanged)
        foreach (var child in report.ChildDiffs.Values)
        {
            var summaryWs = wb.Worksheets.Add($"Child_{child.ChildEntityLogicalName}_Summary".Replace('-', '_'));
            summaryWs.Cell(1, 1).Value = "ChildEntity"; summaryWs.Cell(1, 2).Value = child.ChildEntityLogicalName;
            summaryWs.Cell(2, 1).Value = "ComparedChildFields"; summaryWs.Cell(2, 2).Value = Truncate(string.Join(',', child.ComparedChildFields));
            summaryWs.Cell(3, 1).Value = "SourceChildRecords"; summaryWs.Cell(3, 2).Value = child.TotalSourceChildRecords;
            summaryWs.Cell(4, 1).Value = "TargetChildRecords"; summaryWs.Cell(4, 2).Value = child.TotalTargetChildRecords;
            summaryWs.Cell(5, 1).Value = "ChildMismatches"; summaryWs.Cell(5, 2).Value = child.Mismatches.Count;
            summaryWs.Cell(6, 1).Value = "ChildMultiFieldMismatches"; summaryWs.Cell(6, 2).Value = child.MultiFieldMismatches.Count; // retained in summary, not separate sheet
            summaryWs.Columns().AdjustToContents();

            var childMismatchWs = wb.Worksheets.Add($"Child_{child.ChildEntityLogicalName}_Mismatches".Replace('-', '_'));
            childMismatchWs.Cell(1, 1).Value = "ParentId";
            childMismatchWs.Cell(1, 2).Value = "ChildId";
            childMismatchWs.Cell(1, 3).Value = "FieldName";
            childMismatchWs.Cell(1, 4).Value = "SourceValue";
            childMismatchWs.Cell(1, 5).Value = "TargetValue";
            int cr = 2;
            foreach (var mm in child.Mismatches)
            {
                childMismatchWs.Cell(cr, 1).Value = mm.ParentId.ToString();
                childMismatchWs.Cell(cr, 2).Value = mm.ChildId.ToString();
                childMismatchWs.Cell(cr, 3).Value = mm.FieldName;
                childMismatchWs.Cell(cr, 4).Value = Truncate(mm.SourceValue?.ToString());
                childMismatchWs.Cell(cr, 5).Value = Truncate(mm.TargetValue?.ToString());
                cr++;
            }
            childMismatchWs.Columns().AdjustToContents();

            void WriteChildRecordList(string sheetBase, Dictionary<Guid, List<Dictionary<string, object?>>> dict)
            {
                var ws = wb.Worksheets.Add($"Child_{child.ChildEntityLogicalName}_{sheetBase}".Replace('-', '_'));
                ws.Cell(1, 1).Value = "ParentId";
                ws.Cell(1, 2).Value = "ChildId";
                for (int i = 0; i < child.ComparedChildFields.Count; i++) ws.Cell(1, i + 3).Value = child.ComparedChildFields[i];
                int row = 2;
                foreach (var parentPair in dict)
                {
                    foreach (var rec in parentPair.Value)
                    {
                        rec.TryGetValue(child.ChildEntityLogicalName + "id", out var cid);
                        ws.Cell(row, 1).Value = parentPair.Key.ToString();
                        ws.Cell(row, 2).Value = cid?.ToString();
                        for (int i = 0; i < child.ComparedChildFields.Count; i++)
                        {
                            rec.TryGetValue(child.ComparedChildFields[i], out var val);
                            ws.Cell(row, i + 3).Value = Truncate(val?.ToString());
                        }
                        row++;
                    }
                }
                ws.Columns().AdjustToContents();
            }
            if (child.OnlyInSourceByParent.Count > 0) WriteChildRecordList("OnlyInSource", child.OnlyInSourceByParent);
            if (child.OnlyInTargetByParent.Count > 0) WriteChildRecordList("OnlyInTarget", child.OnlyInTargetByParent);
        }

        using var ms = new System.IO.MemoryStream();
        wb.SaveAs(ms);
        return ms.ToArray();
    }

    private static string GenerateBlobName(string entityLogicalName)
    {
        var safeEntity = new string(entityLogicalName.Select(ch => char.IsLetterOrDigit(ch) ? ch : '_').ToArray());
        return $"comparison-{safeEntity}-{DateTime.UtcNow:yyyyMMddHHmmss}.xlsx";
    }

    private static async Task<string> UploadExcelWithFlexibleAuthAsync(byte[] bytes, string blobName, string? accountUrl = null, string? containerName = null, TokenCredential? credential = null, string? connectionString = null, string? containerSasUrl = null)
    {
        BlobContainerClient containerClient;
        if (!string.IsNullOrWhiteSpace(containerSasUrl))
        {
            var raw = containerSasUrl.Trim();
            // Provide clearer validation so caller knows required format
            if (!raw.StartsWith("http://", StringComparison.OrdinalIgnoreCase) && !raw.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
            {
                throw new ArgumentException("StorageContainerSasUrl must be a full https URL to the container, e.g. https://<account>.blob.core.windows.net/<container>?<sas>");
            }
            if (!Uri.TryCreate(raw, UriKind.Absolute, out var sasUri))
            {
                throw new ArgumentException("StorageContainerSasUrl is not a valid absolute URI after trimming.");
            }
            if (string.IsNullOrEmpty(sasUri.Query))
            {
                throw new ArgumentException("StorageContainerSasUrl is missing SAS query parameters.");
            }
            containerClient = new BlobContainerClient(sasUri);
        }
        else if (!string.IsNullOrWhiteSpace(connectionString) && !string.IsNullOrWhiteSpace(containerName))
        {
            containerClient = new BlobServiceClient(connectionString).GetBlobContainerClient(containerName);
        }
        else if (!string.IsNullOrWhiteSpace(accountUrl) && credential != null && !string.IsNullOrWhiteSpace(containerName))
        {
            containerClient = new BlobServiceClient(new Uri(accountUrl), credential).GetBlobContainerClient(containerName);
        }
        else
        {
            throw new InvalidOperationException("Insufficient storage parameters provided.");
        }
        if (!string.IsNullOrWhiteSpace(connectionString) || (!string.IsNullOrWhiteSpace(accountUrl) && credential != null))
            await containerClient.CreateIfNotExistsAsync();
        using var ms = new System.IO.MemoryStream(bytes);
        var blob = containerClient.GetBlobClient(blobName);
        await blob.UploadAsync(ms, overwrite: true);
        return blob.Uri.ToString();
    }

    private static async Task<List<Dictionary<string, object?>>> FetchChildRecordsAsync(ILogger logger, string envUrl, string childEntityLogicalName, string parentFieldLogicalName, IEnumerable<string> childFields, string accessToken, bool onlyActive, Guid correlationId)
    {
        var client = new HttpClient();
        client.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", accessToken);
        client.DefaultRequestHeaders.Accept.ParseAdd("application/json");
        client.DefaultRequestHeaders.Add("Prefer", "odata.include-annotations=\"OData.Community.Display.V1.FormattedValue\"");

        string pkField = childEntityLogicalName + "id";
        var selectFields = new List<string> { pkField };
        if (!string.IsNullOrWhiteSpace(parentFieldLogicalName)) selectFields.Add(parentFieldLogicalName);
        selectFields.AddRange(childFields);
        selectFields = selectFields.Where(f => !string.IsNullOrWhiteSpace(f))
                                   .Distinct(StringComparer.OrdinalIgnoreCase)
                                   .Where(f => f.Length < 129) // safety length check
                                   .ToList();

        // Separate potential formatted variants (those ending in 'name' or 'typename') so we can remap later if base field exists
        var rawIntersection = new List<string>(selectFields);
        var formattedVariants = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase); // variant -> base
        foreach (var f in rawIntersection.ToList())
        {
            string? baseCandidate = null;
            if (f.EndsWith("typename", StringComparison.OrdinalIgnoreCase)) baseCandidate = f[..^8];
            else if (f.EndsWith("name", StringComparison.OrdinalIgnoreCase)) baseCandidate = f[..^4];
            if (baseCandidate != null && rawIntersection.Contains(baseCandidate))
            {
                formattedVariants[f] = baseCandidate;
                // Remove pure formatted variant from select list (we will inject value later from annotation)
                selectFields.Remove(f);
            }
        }

        string entitySetName = await GetEntitySetNameAsync(envUrl, childEntityLogicalName, accessToken) ?? (childEntityLogicalName + "s");
        string baseUrl = envUrl.TrimEnd('/') + $"/api/data/v9.2/{entitySetName}";
        string filter = onlyActive ? "&$filter=statecode eq 0" : string.Empty;

        var all = new List<Dictionary<string, object?>>();
        int attempt = 0;
        while (true)
        {
            attempt++;
            bool restartRequested = false;
            string selectClause = string.Join(',', selectFields);
            string url = baseUrl + "?$select=" + selectClause + filter;
            logger.LogDebug("[{CorrelationId}] Child fetch attempt={Attempt} Entity={ChildEntity} FieldCount={FieldCount} URLLength={Len}", correlationId, attempt, childEntityLogicalName, selectFields.Count, url.Length);
            try
            {
                int page = 0;
                while (!string.IsNullOrEmpty(url))
                {
                    page++;
                    using var req = new HttpRequestMessage(HttpMethod.Get, url);
                    using var resp = await client.SendAsync(req);
                    if (!resp.IsSuccessStatusCode)
                    {
                        var content = await resp.Content.ReadAsStringAsync();
                        if ((int)resp.StatusCode == 401)
                        {
                            logger.LogError("[{CorrelationId}] Child 401 Entity={ChildEntity} attempt={Attempt} snippet={Snippet}", correlationId, childEntityLogicalName, attempt, content.Length > 200 ? content[..200] : content);
                            throw new UnauthorizedAccessException($"Unauthorized child fetch for {childEntityLogicalName}");
                        }
                        if ((int)resp.StatusCode == 400)
                        {
                            logger.LogWarning("[{CorrelationId}] Child 400 Entity={ChildEntity} attempt={Attempt} snippet={Snippet}", correlationId, childEntityLogicalName, attempt, content.Length > 240 ? content[..240] : content);
                            var invalid = ExtractInvalidProperties(content);
                            if (invalid.Count > 0)
                            {
                                int removed = 0;
                                foreach (var ip in invalid)
                                {
                                    if (!string.Equals(ip, pkField, StringComparison.OrdinalIgnoreCase) && selectFields.Remove(ip))
                                    {
                                        removed++;
                                        logger.LogWarning("[{CorrelationId}] Removed invalid child field '{Field}' Entity={ChildEntity}", correlationId, ip, childEntityLogicalName);
                                    }
                                }
                                if (removed > 0)
                                {
                                    restartRequested = true;
                                    break; // break paging loop to restart with reduced field set
                                }
                            }
                            if (invalid.Count == 0 && selectFields.Count > 2) // fallback: keep only pk & parent field
                            {
                                logger.LogWarning("[{CorrelationId}] Could not isolate invalid child field; fallback to pk+parent only Entity={ChildEntity}", correlationId, childEntityLogicalName);
                                selectFields = selectFields.Where(f => string.Equals(f, pkField, StringComparison.OrdinalIgnoreCase) || string.Equals(f, parentFieldLogicalName, StringComparison.OrdinalIgnoreCase)).ToList();
                                restartRequested = true;
                                break;
                            }
                        }
                        else
                        {
                            resp.EnsureSuccessStatusCode();
                        }
                    }
                    using var stream = await resp.Content.ReadAsStreamAsync();
                    using var doc = await JsonDocument.ParseAsync(stream);
                    if (doc.RootElement.TryGetProperty("value", out var arr))
                    {
                        foreach (var item in arr.EnumerateArray()) all.Add(ToDictionary(item));
                    }
                    url = doc.RootElement.TryGetProperty("@odata.nextLink", out var nextLink) ? nextLink.GetString() : null;
                }
                if (!restartRequested) break; // success without needing restart
            }
            catch (UnauthorizedAccessException) { throw; }
            catch (Exception ex) when (attempt < 5)
            {
                logger.LogWarning(ex, "[{CorrelationId}] Child fetch attempt={Attempt} transient failure Entity={ChildEntity}", correlationId, attempt, childEntityLogicalName);
            }
            if (!restartRequested || attempt >= 8) break; // stop retrying
        }

        // Inject formatted variant values using annotations
        if (formattedVariants.Count > 0 && all.Count > 0)
        {
            foreach (var rec in all)
            {
                foreach (var kv in formattedVariants)
                {
                    var annotationKey = kv.Value + "@OData.Community.Display.V1.FormattedValue";
                    if (rec.TryGetValue(annotationKey, out var fmt)) rec[kv.Key] = fmt;
                    else if (!rec.ContainsKey(kv.Key)) rec[kv.Key] = null;
                }
            }
        }

        logger.LogInformation("[{CorrelationId}] Child fetch completed Entity={ChildEntity} Records={Count} FinalFieldCount={FieldCount}", correlationId, childEntityLogicalName, all.Count, selectFields.Count);
        return all;

        static List<string> ExtractInvalidProperties(string body)
        {
            var list = new List<string>();
            string marker1 = "Could not find a property named '";
            int idx = body.IndexOf(marker1, StringComparison.OrdinalIgnoreCase);
            while (idx >= 0)
            {
                int start = idx + marker1.Length;
                int end = body.IndexOf("'", start);
                if (end > start)
                {
                    list.Add(body[start..end]);
                    idx = body.IndexOf(marker1, end, StringComparison.OrdinalIgnoreCase);
                }
                else break;
            }
            string marker2 = "Property '";
            idx = body.IndexOf(marker2, StringComparison.OrdinalIgnoreCase);
            while (idx >= 0)
            {
                int start = idx + marker2.Length;
                int end = body.IndexOf("'", start);
                if (end > start)
                {
                    var val = body[start..end];
                    if (!list.Contains(val, StringComparer.OrdinalIgnoreCase)) list.Add(val);
                    idx = body.IndexOf(marker2, end, StringComparison.OrdinalIgnoreCase);
                }
                else break;
            }
            return list.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        }
    }

    // Cache for entity set names to avoid repeated metadata calls
    private static readonly System.Collections.Concurrent.ConcurrentDictionary<string, string> _entitySetNameCache = new(StringComparer.OrdinalIgnoreCase);

    private static async Task<string?> GetEntitySetNameAsync(string envUrl, string logicalName, string accessToken)
    {
        if (string.IsNullOrWhiteSpace(logicalName)) return null;
        if (_entitySetNameCache.TryGetValue(logicalName, out var cached)) return cached;
        try
        {
            using var client = new HttpClient();
            client.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", accessToken);
            client.DefaultRequestHeaders.Accept.ParseAdd("application/json");
            var url = envUrl.TrimEnd('/') + $"/api/data/v9.2/EntityDefinitions(LogicalName='{logicalName}')?$select=EntitySetName";
            using var resp = await client.GetAsync(url);
            if (!resp.IsSuccessStatusCode) return null;
            using var stream = await resp.Content.ReadAsStreamAsync();
            using var doc = await JsonDocument.ParseAsync(stream);
            if (doc.RootElement.TryGetProperty("EntitySetName", out var esn))
            {
                var val = esn.GetString();
                if (!string.IsNullOrWhiteSpace(val))
                {
                    _entitySetNameCache[logicalName] = val;
                    return val;
                }
            }
        }
        catch { }
        return null;
    }

    // Retrieve primary name attribute for entity (cached)
    private static readonly System.Collections.Concurrent.ConcurrentDictionary<string, string> _primaryNameAttributeCache = new(StringComparer.OrdinalIgnoreCase);
    private static async Task<string?> GetPrimaryNameAttributeAsync(string envUrl, string logicalName, string accessToken)
    {
        if (string.IsNullOrWhiteSpace(logicalName)) return null;
        if (_primaryNameAttributeCache.TryGetValue(logicalName, out var cached)) return cached;
        try
        {
            using var client = new HttpClient();
            client.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", accessToken);
            client.DefaultRequestHeaders.Accept.ParseAdd("application/json");
            var url = envUrl.TrimEnd('/') + $"/api/data/v9.2/EntityDefinitions(LogicalName='{logicalName}')?$select=PrimaryNameAttribute";
            using var resp = await client.GetAsync(url);
            if (!resp.IsSuccessStatusCode) return null;
            using var stream = await resp.Content.ReadAsStreamAsync();
            using var doc = await JsonDocument.ParseAsync(stream);
            if (doc.RootElement.TryGetProperty("PrimaryNameAttribute", out var pna))
            {
                var val = pna.GetString();
                if (!string.IsNullOrWhiteSpace(val))
                {
                    _primaryNameAttributeCache[logicalName] = val;
                    return val;
                }
            }
        }
        catch { }
        return null;
    }
}
