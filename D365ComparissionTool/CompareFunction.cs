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
using System.Xml.Linq;
using System.Text.RegularExpressions;
using Microsoft.Azure.Functions.Worker.Extensions.Timer;

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

    // Weekly timer trigger (every Friday at 00:00 UTC)
    [Function("CompareDataverseWeekly")]
    public async Task RunWeekly([TimerTrigger("0 0 9 * * Fri")] TimerInfo timerInfo, FunctionContext context)
    {
        var correlationId = Guid.NewGuid();
        _logger.LogInformation("[CompareDataverseWeekly] Timer fired (Fri 09:00 UTC) CorrelationId={CorrelationId} LastRun={LastRun} NextRun={NextRun}", correlationId, timerInfo?.ScheduleStatus?.Last, timerInfo?.ScheduleStatus?.Next);
        try
        {
            var srcEnv = Environment.GetEnvironmentVariable("SOURCE_ENV_URL") ?? string.Empty;
            var tgtEnv = Environment.GetEnvironmentVariable("TARGET_ENV_URL") ?? string.Empty;
            if (string.IsNullOrWhiteSpace(srcEnv) || string.IsNullOrWhiteSpace(tgtEnv))
            {
                _logger.LogWarning("[{CorrelationId}] SOURCE_ENV_URL or TARGET_ENV_URL missing; skipping weekly comparison", correlationId);
                return;
            }
            // Determine mode (single | list) from env, default list
            var modeEnv = (Environment.GetEnvironmentVariable("COMPARISON_WEEKLY_MODE") ?? "list").Trim().ToLowerInvariant();

            // Parse entities from JSON array env or fallback comma-separated list
            List<string> entities = new();
            var entitiesJson = Environment.GetEnvironmentVariable("COMPARISON_WEEKLY_ENTITIES_JSON");
            if (!string.IsNullOrWhiteSpace(entitiesJson))
            {
                try
                {
                    var parsed = JsonSerializer.Deserialize<List<string>>(entitiesJson);
                    if (parsed != null) entities.AddRange(parsed);
                }
                catch (Exception jex)
                {
                    _logger.LogWarning(jex, "[{CorrelationId}] Failed to parse COMPARISON_WEEKLY_ENTITIES_JSON; falling back to COMPARISON_WEEKLY_ENTITIES", correlationId);
                }
            }
            if (entities.Count == 0)
            {
                var entitiesRaw = Environment.GetEnvironmentVariable("COMPARISON_WEEKLY_ENTITIES") ?? string.Empty;
                if (!string.IsNullOrWhiteSpace(entitiesRaw))
                {
                    entities.AddRange(entitiesRaw.Split(new[] { ',', ';', ' ' }, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries));
                }
            }
            entities = entities.Select(e => e.Trim()).Where(e => !string.IsNullOrWhiteSpace(e)).Distinct(StringComparer.OrdinalIgnoreCase).ToList();

            var singleEntityEnv = Environment.GetEnvironmentVariable("COMPARISON_WEEKLY_ENTITY")?.Trim();
            bool singleMode = modeEnv == "single" || (!string.IsNullOrWhiteSpace(singleEntityEnv)) || (entities.Count == 1);

            var request = new ComparisonRequest
            {
                Mode = singleMode ? "single" : "list",
                SourceEnvUrl = srcEnv,
                TargetEnvUrl = tgtEnv,
                AutoDiscoverSubgrids = false
            };

            if (singleMode)
            {
                // Determine entity logical name priority: explicit single env variable > first of list
                string? entityLogical = singleEntityEnv;
                if (string.IsNullOrWhiteSpace(entityLogical)) entityLogical = entities.FirstOrDefault();
                if (string.IsNullOrWhiteSpace(entityLogical))
                {
                    _logger.LogWarning("[{CorrelationId}] No entity provided for single weekly comparison; aborting", correlationId);
                    return;
                }
                request.EntityLogicalName = entityLogical;
                _logger.LogInformation("[{CorrelationId}] Weekly single entity comparison Entity={Entity}", correlationId, entityLogical);
            }
            else
            {
                if (entities.Count == 0)
                {
                    _logger.LogWarning("[{CorrelationId}] No entities specified for weekly list comparison; skipping", correlationId);
                    return;
                }
                request.Entities = entities;
                _logger.LogInformation("[{CorrelationId}] Weekly list entity comparison Count={Count} Entities={Entities}", correlationId, entities.Count, string.Join(',', entities));
            }

            if (bool.TryParse(Environment.GetEnvironmentVariable("COMPARISON_WEEKLY_AUTODISCOVER_SUBGRIDS"), out var ad)) request.AutoDiscoverSubgrids = ad;
            request.SourceBearerToken = Environment.GetEnvironmentVariable("SOURCE_BEARER_TOKEN");
            request.TargetBearerToken = Environment.GetEnvironmentVariable("TARGET_BEARER_TOKEN");
            request.StorageAccountUrl = Environment.GetEnvironmentVariable("COMPARISON_STORAGE_URL");
            request.OutputContainerName = Environment.GetEnvironmentVariable("COMPARISON_OUTPUT_CONTAINER");
            request.StorageConnectionString = Environment.GetEnvironmentVariable("COMPARISON_STORAGE_CONNECTION_STRING");
            request.StorageContainerSasUrl = Environment.GetEnvironmentVariable("COMPARISON_STORAGE_CONTAINER_SAS_URL");

            var credential = new DefaultAzureCredential();
            string sourceToken = !string.IsNullOrWhiteSpace(request.SourceBearerToken) ? request.SourceBearerToken.Trim() : await AcquireSourceTokenAsync(request, credential);

            // Determine multi-target set (always compare source against each target environment)
            var targetEnvList = new List<string>();
            // Parse COMPARISON_WEEKLY_TARGETS (JSON array or comma-separated) if present
            var multiTargetsRaw = Environment.GetEnvironmentVariable("COMPARISON_WEEKLY_TARGETS");
            if (!string.IsNullOrWhiteSpace(multiTargetsRaw))
            {
                bool parsed = false;
                try
                {
                    var arr = JsonSerializer.Deserialize<List<string>>(multiTargetsRaw);
                    if (arr != null && arr.Count > 0)
                    {
                        targetEnvList.AddRange(arr.Where(a => !string.IsNullOrWhiteSpace(a))); parsed = true;
                    }
                }
                catch { }
                if (!parsed)
                {
                    targetEnvList.AddRange(multiTargetsRaw.Split(new[] { ',', ';', ' ' }, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries));
                }
            }
            if (targetEnvList.Count == 0)
            {
                // Fallback: use TARGET_ENV_URL env plus predefined list if source is mash and not already provided
                targetEnvList.Add(tgtEnv.Trim());
                if (srcEnv.Trim().StartsWith("https://mash.crm.dynamics.com", StringComparison.OrdinalIgnoreCase))
                {
                    var predefined = new[]
                    {
                        "https://mashppe.crm.dynamics.com/",
                        "https://mashtest.crm.dynamics.com/",
                        "https://mashrb.crm.dynamics.com/",
                        "https://mashdemo.crm.dynamics.com/"
                    };
                    foreach (var p in predefined)
                        if (!targetEnvList.Contains(p, StringComparer.OrdinalIgnoreCase)) targetEnvList.Add(p);
                }
            }
            targetEnvList = targetEnvList.Select(e => e.Trim()).Where(e => !string.IsNullOrWhiteSpace(e)).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
            _logger.LogInformation("[{CorrelationId}] Weekly comparison target environments Count={Count} Targets={Targets}", correlationId, targetEnvList.Count, string.Join(',', targetEnvList));

            foreach (var targetEnv in targetEnvList)
            {
                request.TargetEnvUrl = targetEnv; // set target for this iteration
                string targetToken = !string.IsNullOrWhiteSpace(request.TargetBearerToken) ? request.TargetBearerToken.Trim() : await AcquireTargetTokenAsync(request, credential);
                _logger.LogInformation("[{CorrelationId}] Processing target {TargetEnv} Mode={Mode}", correlationId, targetEnv, request.Mode);
                if (singleMode)
                {
                    try
                    {
                        var singleReport = await CompareSingleEntityAsync(request, sourceToken, targetToken, credential, correlationId, skipExcelUpload: false);
                        _logger.LogInformation("[{CorrelationId}] Weekly single entity comparison finished Target={TargetEnv} Entity={Entity} Excel={Excel}", correlationId, targetEnv, request.EntityLogicalName, singleReport.ExcelBlobUrl);
                    }
                    catch (Exception exSingle)
                    {
                        _logger.LogError(exSingle, "[{CorrelationId}] Weekly single entity comparison failed Target={TargetEnv} Entity={Entity}", correlationId, targetEnv, request.EntityLogicalName);
                    }
                }
                else
                {
                    var results = new List<(string Entity, DiffReport Report)>();
                    foreach (var entity in request.Entities)
                    {
                        request.EntityLogicalName = entity;
                        var report = await CompareSingleEntityAsync(request, sourceToken, targetToken, credential, correlationId, skipExcelUpload: true);
                        results.Add((entity, report));
                    }
                    try
                    {
                        byte[] excelBytes = BuildMultiEntityExcel(results, request.SourceEnvUrl, targetEnv, DateTime.UtcNow);
                        string blobName = GenerateBlobName(request.SourceEnvUrl, targetEnv, "multi", isCombined: true);
                        string containerName = request.OutputContainerName ?? "comparissiontooloutput";
                        string storageUrl = request.StorageAccountUrl ?? string.Empty;
                        string? uploadedUrl = null;
                        if (!string.IsNullOrWhiteSpace(request.StorageContainerSasUrl)) uploadedUrl = await UploadExcelWithFlexibleAuthAsync(excelBytes, blobName, containerSasUrl: request.StorageContainerSasUrl);
                        else if (!string.IsNullOrWhiteSpace(request.StorageConnectionString)) uploadedUrl = await UploadExcelWithFlexibleAuthAsync(excelBytes, blobName, connectionString: request.StorageConnectionString, containerName: containerName);
                        else if (!string.IsNullOrWhiteSpace(storageUrl)) uploadedUrl = await UploadExcelWithFlexibleAuthAsync(excelBytes, blobName, accountUrl: storageUrl, containerName: containerName, credential: credential);
                        if (!string.IsNullOrWhiteSpace(uploadedUrl)) _logger.LogInformation("[{CorrelationId}] Weekly list comparison uploaded Excel={BlobUrl} Target={TargetEnv} Entities={EntityCount}", correlationId, uploadedUrl, targetEnv, results.Count);
                        else _logger.LogWarning("[{CorrelationId}] Weekly list comparison Excel not uploaded (no storage config) Target={TargetEnv}", correlationId, targetEnv);
                    }
                    catch (Exception upEx)
                    {
                        _logger.LogError(upEx, "[{CorrelationId}] Weekly list Excel generation/upload failed Target={TargetEnv}", correlationId, targetEnv);
                    }
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "[CompareDataverseWeekly] Unhandled error CorrelationId={CorrelationId}", correlationId);
        }
    }

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
                    // Skip individual Excel uploads; aggregate later
                    var report = await CompareSingleEntityAsync(request, singleSourceToken, singleTargetToken, credential, correlationId, skipExcelUpload: true);
                    results.Add((entity, report));
                    request.EntityLogicalName = original; // restore
                }
                // Build combined Excel if storage settings provided
                string combinedBlobUrl = null;
                try
                {
                    string storageUrl = request.StorageAccountUrl ?? Environment.GetEnvironmentVariable("COMPARISON_STORAGE_URL") ?? string.Empty;
                    string containerName = request.OutputContainerName ?? Environment.GetEnvironmentVariable("COMPARISON_OUTPUT_CONTAINER") ?? "comparissiontooloutput";
                    byte[] excelBytes = BuildMultiEntityExcel(results, request.SourceEnvUrl, request.TargetEnvUrl, DateTime.UtcNow);
                    _logger.LogInformation("[{CorrelationId}] Combined Excel size bytes={Size}", correlationId, excelBytes.Length);
                    if (!string.IsNullOrWhiteSpace(request.StorageContainerSasUrl))
                    {
                        combinedBlobUrl = await UploadExcelWithFlexibleAuthAsync(excelBytes, GenerateBlobName(request.SourceEnvUrl, request.TargetEnvUrl, "multi", isCombined: true), containerSasUrl: request.StorageContainerSasUrl);
                    }
                    else if (!string.IsNullOrWhiteSpace(request.StorageConnectionString))
                    {
                        combinedBlobUrl = await UploadExcelWithFlexibleAuthAsync(excelBytes, GenerateBlobName(request.SourceEnvUrl, request.TargetEnvUrl, "multi", isCombined: true), connectionString: request.StorageConnectionString, containerName: containerName);
                    }
                    else if (!string.IsNullOrWhiteSpace(storageUrl))
                    {
                        var credential2 = new DefaultAzureCredential();
                        combinedBlobUrl = await UploadExcelWithFlexibleAuthAsync(excelBytes, GenerateBlobName(request.SourceEnvUrl, request.TargetEnvUrl, "multi", isCombined: true), accountUrl: storageUrl, containerName: containerName, credential: credential2);
                    }
                    else
                    {
                        _logger.LogWarning("[{CorrelationId}] No storage configuration provided; combined Excel not uploaded", correlationId);
                    }
                }
                catch (Exception exUp)
                {
                    _logger.LogError(exUp, "[{CorrelationId}] Combined Excel upload failed", correlationId);
                }

                var payload = new
                {
                    Mode = mode,
                    CombinedExcelBlobUrl = combinedBlobUrl,
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

    private async Task<DiffReport> CompareSingleEntityAsync(ComparisonRequest request, string sourceToken, string targetToken, DefaultAzureCredential credential, Guid correlationId, bool skipExcelUpload = false)
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

        if (!skipExcelUpload)
        {
            // Excel upload (optional - single entity mode)
            string storageUrl = request.StorageAccountUrl ?? Environment.GetEnvironmentVariable("COMPARISON_STORAGE_URL") ?? string.Empty;
            string containerName = request.OutputContainerName ?? Environment.GetEnvironmentVariable("COMPARISON_OUTPUT_CONTAINER") ?? "comparissiontooloutput";
            if (!string.IsNullOrWhiteSpace(request.StorageContainerSasUrl))
            {
                _logger.LogInformation("[{CorrelationId}] Using SAS URL for upload", correlationId);
                try
                {
                    string blobName = GenerateBlobName(request.SourceEnvUrl, request.TargetEnvUrl, request.EntityLogicalName);
                    var excelBytes = BuildExcel(report, request.EntityLogicalName, request.SourceEnvUrl, request.TargetEnvUrl, DateTime.UtcNow);
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
                    string blobName = GenerateBlobName(request.SourceEnvUrl, request.TargetEnvUrl, request.EntityLogicalName);
                    var excelBytes = BuildExcel(report, request.EntityLogicalName, request.SourceEnvUrl, request.TargetEnvUrl, DateTime.UtcNow);
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
                    string blobName = GenerateBlobName(request.SourceEnvUrl, request.TargetEnvUrl, request.EntityLogicalName);
                    var excelBytes = BuildExcel(report, request.EntityLogicalName, request.SourceEnvUrl, request.TargetEnvUrl, DateTime.UtcNow);
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
        }

        _logger.LogInformation("[{CorrelationId}] Comparison completed Entity={Entity} ExcelUrl={ExcelUrl}", correlationId, request.EntityLogicalName, report.ExcelBlobUrl ?? "(none)");
        return report;
    }

    // Build a combined multi-entity Excel workbook (one run, multiple entities)
    private static byte[] BuildMultiEntityExcel(List<(string Entity, DiffReport Report)> results, string sourceEnvUrl, string targetEnvUrl, DateTime comparisonTimestampUtc)
    {
        const int ExcelMaxCellLength = 32767;
        string Truncate(string? v) => string.IsNullOrEmpty(v) ? v ?? string.Empty : (v.Length <= ExcelMaxCellLength ? v : v.Substring(0, ExcelMaxCellLength - 15) + "...(truncated)");
        string BuildRecordUrl(string envUrl, string etn, Guid id) => string.IsNullOrWhiteSpace(envUrl) || string.IsNullOrWhiteSpace(etn) || id == Guid.Empty ? string.Empty : envUrl.TrimEnd('/') + $"/main.aspx?etn={etn}&id={id}&pagetype=entityrecord";

        using var wb = new XLWorkbook();

        // Summary sheet
        var summary = wb.Worksheets.Add("Summary");
        summary.Cell(1,1).Value = "Source Env";
        summary.Cell(1,2).Value = "Target Env";
        summary.Cell(1,3).Value = "Comparission Date and Time";
        summary.Cell(1,4).Value = "Entity";
        summary.Cell(1,5).Value = "Compared Fields";
        summary.Cell(1,6).Value = "Matching Records";
        summary.Cell(1,7).Value = "Mismatches";
        summary.Cell(1,8).Value = "Only In Source";
        summary.Cell(1,9).Value = "Only In Target";
        summary.Cell(1,10).Value = "Notes";
        int sr = 2;
        foreach (var (Entity, Report) in results)
        {
            summary.Cell(sr,1).Value = sourceEnvUrl;
            summary.Cell(sr,2).Value = targetEnvUrl;
            summary.Cell(sr,3).Value = comparisonTimestampUtc.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss");
            summary.Cell(sr,4).Value = Entity;
            summary.Cell(sr,5).Value = Truncate(string.Join(',', Report.ComparedFields));
            summary.Cell(sr,6).Value = Report.MatchingRecordIds.Count;
            summary.Cell(sr,7).Value = Report.Mismatches.Count;
            summary.Cell(sr,8).Value = Report.OnlyInSource.Count;
            summary.Cell(sr,9).Value = Report.OnlyInTarget.Count;
            summary.Cell(sr,10).Value = Truncate(Report.Notes);
            sr++;
        }
        summary.Columns().AdjustToContents();

        // Aggregated MatchingRecords sheet
        var matchWs = wb.Worksheets.Add("MatchingRecords");
        matchWs.Cell(1,1).Value = "Entity";
        matchWs.Cell(1,2).Value = "RecordId";
        matchWs.Cell(1,3).Value = "SourceRecordUrl";
        matchWs.Cell(1,4).Value = "TargetRecordUrl";
        int mr = 2;
        foreach (var (Entity, Report) in results)
        {
            foreach (var id in Report.MatchingRecordIds)
            {
                matchWs.Cell(mr,1).Value = Entity;
                matchWs.Cell(mr,2).Value = id.ToString();
                matchWs.Cell(mr,3).Value = BuildRecordUrl(sourceEnvUrl, Entity, id);
                matchWs.Cell(mr,4).Value = BuildRecordUrl(targetEnvUrl, Entity, id);
                mr++;
            }
        }
        matchWs.Columns().AdjustToContents();

        // Aggregated Mismatches sheet
        var mismatchWs = wb.Worksheets.Add("Mismatches");
        mismatchWs.Cell(1,1).Value = "Entity";
        mismatchWs.Cell(1,2).Value = "RecordId";
        mismatchWs.Cell(1,3).Value = "SourceRecordUrl";
        mismatchWs.Cell(1,4).Value = "TargetRecordUrl";
        mismatchWs.Cell(1,5).Value = "FieldName";
        mismatchWs.Cell(1,6).Value = "SourceValue";
        mismatchWs.Cell(1,7).Value = "TargetValue";
        int mmr = 2;
        foreach (var (Entity, Report) in results)
        {
            foreach (var m in Report.Mismatches)
            {
                mismatchWs.Cell(mmr,1).Value = Entity;
                mismatchWs.Cell(mmr,2).Value = m.RecordId.ToString();
                mismatchWs.Cell(mmr,3).Value = BuildRecordUrl(sourceEnvUrl, Entity, m.RecordId);
                mismatchWs.Cell(mmr,4).Value = BuildRecordUrl(targetEnvUrl, Entity, m.RecordId);
                mismatchWs.Cell(mmr,5).Value = m.FieldName;
                mismatchWs.Cell(mmr,6).Value = Truncate(m.SourceValue?.ToString());
                mismatchWs.Cell(mmr,7).Value = Truncate(m.TargetValue?.ToString());
                mmr++;
            }
        }
        mismatchWs.Columns().AdjustToContents();

        // Aggregated OnlyInSource sheet
        var onlySrcWs = wb.Worksheets.Add("OnlyInSource");
        onlySrcWs.Cell(1,1).Value = "Entity";
        onlySrcWs.Cell(1,2).Value = "RecordId";
        onlySrcWs.Cell(1,3).Value = "SourceRecordUrl";
        int osr = 2; 
        foreach (var (Entity, Report) in results)
        {
            string pk = Entity + "id";
            foreach (var rec in Report.OnlyInSource)
            {
                if (rec.TryGetValue(pk, out var idVal) && Guid.TryParse(idVal?.ToString(), out var gid))
                {
                    onlySrcWs.Cell(osr,1).Value = Entity;
                    onlySrcWs.Cell(osr,2).Value = gid.ToString();
                    onlySrcWs.Cell(osr,3).Value = BuildRecordUrl(sourceEnvUrl, Entity, gid);
                    osr++;
                }
            }
        }
        onlySrcWs.Columns().AdjustToContents();

        // Aggregated OnlyInTarget sheet
        var onlyTgtWs = wb.Worksheets.Add("OnlyInTarget");
        onlyTgtWs.Cell(1,1).Value = "Entity";
        onlyTgtWs.Cell(1,2).Value = "RecordId";
        onlyTgtWs.Cell(1,3).Value = "TargetRecordUrl";
        int otr = 2; 
        foreach (var (Entity, Report) in results)
        {
            string pk = Entity + "id";
            foreach (var rec in Report.OnlyInTarget)
            {
                if (rec.TryGetValue(pk, out var idVal) && Guid.TryParse(idVal?.ToString(), out var gid))
                {
                    onlyTgtWs.Cell(otr,1).Value = Entity;
                    onlyTgtWs.Cell(otr,2).Value = gid.ToString();
                    onlyTgtWs.Cell(otr,3).Value = BuildRecordUrl(targetEnvUrl, Entity, gid);
                    otr++;
                }
            }
        }
        onlyTgtWs.Columns().AdjustToContents();

        // Child diff sheets remain per entity to avoid overly large aggregation
        foreach (var (Entity, Report) in results)
        {
            foreach (var child in Report.ChildDiffs.Values)
            {
                var cSummary = wb.Worksheets.Add($"{Entity}_Child_{child.ChildEntityLogicalName}_Summary".Replace('-', '_'));
                cSummary.Cell(1,1).Value = "ParentEntity"; cSummary.Cell(1,2).Value = Entity;
                cSummary.Cell(2,1).Value = "ChildEntity"; cSummary.Cell(2,2).Value = child.ChildEntityLogicalName;
                cSummary.Cell(3,1).Value = "ComparedChildFields"; cSummary.Cell(3,2).Value = Truncate(string.Join(',', child.ComparedChildFields));
                cSummary.Cell(4,1).Value = "SourceChildRecords"; cSummary.Cell(4,2).Value = child.TotalSourceChildRecords;
                cSummary.Cell(5,1).Value = "TargetChildRecords"; cSummary.Cell(5,2).Value = child.TotalTargetChildRecords;
                cSummary.Cell(6,1).Value = "ChildMismatches"; cSummary.Cell(6,2).Value = child.Mismatches.Count;
                cSummary.Cell(7,1).Value = "ChildMultiFieldMismatches"; cSummary.Cell(7,2).Value = child.MultiFieldMismatches.Count;
                cSummary.Columns().AdjustToContents();

                var cMismatch = wb.Worksheets.Add($"{Entity}_Child_{child.ChildEntityLogicalName}_Mismatches".Replace('-', '_'));
                cMismatch.Cell(1,1).Value = "ParentId";
                cMismatch.Cell(1,2).Value = "ChildId";
                cMismatch.Cell(1,3).Value = "SourceChildUrl";
                cMismatch.Cell(1,4).Value = "TargetChildUrl";
                cMismatch.Cell(1,5).Value = "FieldName";
                cMismatch.Cell(1,6).Value = "SourceValue";
                cMismatch.Cell(1,7).Value = "TargetValue";
                int cr = 2;
                foreach (var mm in child.Mismatches)
                {
                    cMismatch.Cell(cr,1).Value = mm.ParentId.ToString();
                    cMismatch.Cell(cr,2).Value = mm.ChildId.ToString();
                    cMismatch.Cell(cr,3).Value = BuildRecordUrl(sourceEnvUrl, child.ChildEntityLogicalName, mm.ChildId);
                    cMismatch.Cell(cr,4).Value = BuildRecordUrl(targetEnvUrl, child.ChildEntityLogicalName, mm.ChildId);
                    cMismatch.Cell(cr,5).Value = mm.FieldName;
                    cMismatch.Cell(cr,6).Value = Truncate(mm.SourceValue?.ToString());
                    cMismatch.Cell(cr,7).Value = Truncate(mm.TargetValue?.ToString());
                    cr++;
                }
                cMismatch.Columns().AdjustToContents();
            }
        }

        using var ms = new System.IO.MemoryStream();
        wb.SaveAs(ms);
        return ms.ToArray();
    }

    // Build combined workbook for multiple entities (list mode) with prefixed sheet names per entity
    private static byte[] BuildCombinedExcel(List<(string EntityLogicalName, DiffReport Report)> results, string sourceEnvUrl, string targetEnvUrl, DateTime utcNow)
    {
        using var wb = new XLWorkbook();
        // Global summary sheet
        var summary = wb.Worksheets.Add("Summary");
        summary.Cell(1,1).Value = "Entity";
        summary.Cell(1,2).Value = "ComparedFields";
        summary.Cell(1,3).Value = "MatchingRecords";
        summary.Cell(1,4).Value = "Mismatches";
        summary.Cell(1,5).Value = "OnlyInSource";
        summary.Cell(1,6).Value = "OnlyInTarget";
        summary.Cell(1,7).Value = "Notes";
        int sr = 2;
        foreach (var item in results)
        {
            summary.Cell(sr,1).Value = item.EntityLogicalName;
            summary.Cell(sr,2).Value = string.Join(',', item.Report.ComparedFields);
            summary.Cell(sr,3).Value = item.Report.MatchingRecordIds.Count;
            summary.Cell(sr,4).Value = item.Report.Mismatches.Count;
            summary.Cell(sr,5).Value = item.Report.OnlyInSource.Count;
            summary.Cell(sr,6).Value = item.Report.OnlyInTarget.Count;
            summary.Cell(sr,7).Value = item.Report.Notes;
            sr++;
        }
        summary.Columns().AdjustToContents();

        // Per-entity detailed sheets (similar to BuildExcel but prefixed)
        foreach (var item in results)
        {
            var entity = item.EntityLogicalName;
            var report = item.Report;
            string prefix = entity + "_"; // sheet name prefix
            // Matching records
            var matchWs = wb.Worksheets.Add(prefix + "MatchingRecords");
            matchWs.Cell(1,1).Value = "Entity";
            matchWs.Cell(1,2).Value = "RecordId";
            matchWs.Cell(1,3).Value = "SourceRecordUrl";
            matchWs.Cell(1,4).Value = "TargetRecordUrl";
            for (int i=0;i<report.MatchingRecordIds.Count;i++)
            {
                var id = report.MatchingRecordIds[i];
                matchWs.Cell(i+2,1).Value = entity;
                matchWs.Cell(i+2,2).Value = id.ToString();
                matchWs.Cell(i+2,3).Value = sourceEnvUrl.TrimEnd('/') + $"/main.aspx?etn={entity}&id={id}&pagetype=entityrecord";
                matchWs.Cell(i+2,4).Value = targetEnvUrl.TrimEnd('/') + $"/main.aspx?etn={entity}&id={id}&pagetype=entityrecord";
            }
            matchWs.Columns().AdjustToContents();

            var mismatchWs = wb.Worksheets.Add(prefix + "Mismatches");
            mismatchWs.Cell(1,1).Value = "Entity";
            mismatchWs.Cell(1,2).Value = "RecordId";
            mismatchWs.Cell(1,3).Value = "SourceRecordUrl";
            mismatchWs.Cell(1,4).Value = "TargetRecordUrl";
            mismatchWs.Cell(1,5).Value = "FieldName";
            mismatchWs.Cell(1,6).Value = "SourceValue";
            mismatchWs.Cell(1,7).Value = "TargetValue";
            int r = 2;
            foreach (var mm in report.Mismatches)
            {
                mismatchWs.Cell(r,1).Value = entity;
                mismatchWs.Cell(r,2).Value = mm.RecordId.ToString();
                mismatchWs.Cell(r,3).Value = sourceEnvUrl.TrimEnd('/') + $"/main.aspx?etn={entity}&id={mm.RecordId}&pagetype=entityrecord";
                mismatchWs.Cell(r,4).Value = targetEnvUrl.TrimEnd('/') + $"/main.aspx?etn={entity}&id={mm.RecordId}&pagetype=entityrecord";
                mismatchWs.Cell(r,5).Value = mm.FieldName;
                mismatchWs.Cell(r,6).Value = mm.SourceValue?.ToString();
                mismatchWs.Cell(r,7).Value = mm.TargetValue?.ToString();
                r++;
            }
            mismatchWs.Columns().AdjustToContents();

            void WriteRecordList(string sheetSuffix, List<Dictionary<string, object?>> records, bool isSource)
            {
                var ws = wb.Worksheets.Add(prefix + sheetSuffix);
                ws.Cell(1,1).Value = "Entity";
                ws.Cell(1,2).Value = "RecordId";
                ws.Cell(1,3).Value = isSource ? "SourceRecordUrl" : "TargetRecordUrl";
                string pk = entity + "id";
                int row = 2;
                foreach (var rec in records)
                {
                    if (rec.TryGetValue(pk, out var idVal) && Guid.TryParse(idVal?.ToString(), out var gid))
                    {
                        ws.Cell(row,1).Value = entity;
                        ws.Cell(row,2).Value = gid.ToString();
                        ws.Cell(row,3).Value = (isSource?sourceEnvUrl:targetEnvUrl).TrimEnd('/') + $"/main.aspx?etn={entity}&id={gid}&pagetype=entityrecord";
                        row++;
                    }
                }
                ws.Columns().AdjustToContents();
            }
            if (report.OnlyInSource.Count>0) WriteRecordList("OnlyInSource", report.OnlyInSource, true);
            if (report.OnlyInTarget.Count>0) WriteRecordList("OnlyInTarget", report.OnlyInTarget, false);
        }

        using var ms = new System.IO.MemoryStream();
        wb.SaveAs(ms);
        return ms.ToArray();
    }

    // Discover one-to-many relationships for the parent entity and populate SubgridRelationships if empty
    private async Task AutoDiscoverSubgridRelationshipsAsync(ComparisonRequest request, string sourceToken, Guid correlationId)
    {
        if (string.IsNullOrWhiteSpace(request.EntityLogicalName)) return;
        try
        {
            _logger.LogInformation("[{CorrelationId}] Auto-discovering subgrid relationships from forms for Entity={Entity}", correlationId, request.EntityLogicalName);

            // Step 1: Get set of relationship schema names actually used in subgrid controls on active forms
            var relationshipSchemaNames = await ExtractSubgridRelationshipSchemaNamesFromFormsAsync(request.SourceEnvUrl, request.EntityLogicalName, sourceToken, correlationId);
            if (relationshipSchemaNames.Count == 0)
            {
                _logger.LogInformation("[{CorrelationId}] No subgrid controls found on forms for Entity={Entity}; skipping child comparison", correlationId, request.EntityLogicalName);
                return;
            }
            _logger.LogInformation("[{CorrelationId}] Form subgrid relationship schema names Count={Count} Names={Names}", correlationId, relationshipSchemaNames.Count, string.Join(',', relationshipSchemaNames));

            // Step 2: Retrieve one-to-many relationship metadata and filter to only those used on forms
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
                    string schemaName = rel.TryGetProperty("SchemaName", out var sn) ? sn.GetString() ?? string.Empty : string.Empty;
                    if (string.IsNullOrWhiteSpace(schemaName)) continue;
                    // Only retain relationships whose schema name appears in a subgrid control
                    if (!relationshipSchemaNames.Contains(schemaName, StringComparer.OrdinalIgnoreCase)) continue;
                    string referencingEntity = rel.TryGetProperty("ReferencingEntity", out var re) ? re.GetString() ?? string.Empty : string.Empty;
                    string referencingAttribute = rel.TryGetProperty("ReferencingAttribute", out var ra) ? ra.GetString() ?? string.Empty : string.Empty;
                    if (string.IsNullOrWhiteSpace(referencingEntity) || string.IsNullOrWhiteSpace(referencingAttribute)) continue;
                    if (referencingEntity.StartsWith("msdyn_", StringComparison.OrdinalIgnoreCase)) continue; // skip large msdyn children
                    list.Add(new SubgridRelationRequest
                    {
                        ChildEntityLogicalName = referencingEntity,
                        ChildParentFieldLogicalName = referencingAttribute,
                        OnlyActiveChildren = true
                    });
                }
            }
            request.SubgridRelationships = list
                .GroupBy(l => (l.ChildEntityLogicalName.ToLowerInvariant(), l.ChildParentFieldLogicalName.ToLowerInvariant()))
                .Select(g => g.First()).ToList();
            _logger.LogInformation("[{CorrelationId}] Auto-discovered form-based SubgridRelationships Count={Count}", correlationId, request.SubgridRelationships.Count);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "[{CorrelationId}] Auto-discovery of subgrids failed", correlationId);
        }
    }

    // Extract relationship schema names referenced by subgrid controls in main forms for given entity
    private async Task<List<string>> ExtractSubgridRelationshipSchemaNamesFromFormsAsync(string envUrl, string entityLogicalName, string accessToken, Guid correlationId)
    {
        var result = new List<string>();
        try
        {
            using var client = new HttpClient();
            client.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", accessToken);
            client.DefaultRequestHeaders.Accept.ParseAdd("application/json");
            // systemform type 2 = main form. objecttypecode filter by logical name.
            // We only need formxml column.
            string url = envUrl.TrimEnd('/') + $"/api/data/v9.2/systemforms?$select=formxml&$filter=objecttypecode eq '{entityLogicalName}' and type eq 2";
            using var resp = await client.GetAsync(url);
            if (!resp.IsSuccessStatusCode)
            {
                var body = await resp.Content.ReadAsStringAsync();
                _logger.LogDebug("[{CorrelationId}] systemforms query failed status={Status} snippet={Snippet}", correlationId, resp.StatusCode, body.Length > 180 ? body[..180] : body);
                return result;
            }
            using var stream = await resp.Content.ReadAsStreamAsync();
            using var doc = await JsonDocument.ParseAsync(stream);
            if (!doc.RootElement.TryGetProperty("value", out var arr)) return result;
            foreach (var form in arr.EnumerateArray())
            {
                var formXml = form.TryGetProperty("formxml", out var fx) ? fx.GetString() : null;
                if (string.IsNullOrWhiteSpace(formXml)) continue;
                try
                {
                    var xdoc = XDocument.Parse(formXml);
                    // Subgrid controls typically have <control ... controltype="subgrid" ...> or type="subgrid". Relationship name stored in relationshipname attribute or parameter node.
                    var controls = xdoc.Descendants().Where(e => string.Equals(e.Name.LocalName, "control", StringComparison.OrdinalIgnoreCase));
                    foreach (var c in controls)
                    {
                        var typeAttr = c.Attribute("controltype") ?? c.Attribute("type");
                        if (typeAttr == null) continue;
                        if (!string.Equals(typeAttr.Value, "subgrid", StringComparison.OrdinalIgnoreCase)) continue;
                        var relAttr = c.Attribute("relationshipname");
                        if (relAttr != null && !string.IsNullOrWhiteSpace(relAttr.Value))
                        {
                            result.Add(relAttr.Value.Trim());
                            continue;
                        }
                        // Sometimes relationship name nested inside parameters node <parameters><relationshipname>...</relationshipname></parameters>
                        var relParam = c.Descendants().FirstOrDefault(d => string.Equals(d.Name.LocalName, "relationshipname", StringComparison.OrdinalIgnoreCase));
                        if (relParam != null && !string.IsNullOrWhiteSpace(relParam.Value)) result.Add(relParam.Value.Trim());
                    }
                }
                catch (Exception xmlEx)
                {
                    _logger.LogDebug(xmlEx, "[{CorrelationId}] Failed parsing form XML for entity={Entity}", correlationId, entityLogicalName);
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogDebug(ex, "[{CorrelationId}] ExtractSubgridRelationshipSchemaNamesFromFormsAsync failed entity={Entity}", correlationId, entityLogicalName);
        }
        return result.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
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
        string? url = envUrl.TrimEnd('/') + $"/api/data/v9.2/EntityDefinitions(LogicalName='{entityLogicalName}')/Attributes?$select=LogicalName";
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
        // Resolve real entity set name from metadata to avoid 404 due to incorrect naive pluralization
        string entitySetName = await GetEntitySetNameAsync(envUrl, entityLogicalName, accessToken) ?? (entityLogicalName + "s");
        if (entitySetName == entityLogicalName + "s")
        {
            logger.LogDebug("[{CorrelationId}] Using fallback pluralized entity set name '{EntitySetName}' for entity '{EntityLogicalName}'", correlationId, entitySetName, entityLogicalName);
        }
        else
        {
            logger.LogDebug("[{CorrelationId}] Resolved entity set name '{EntitySetName}' for entity '{EntityLogicalName}'", correlationId, entitySetName, entityLogicalName);
        }
        string baseUrl = envUrl.TrimEnd('/') + $"/api/data/v9.2/{entitySetName}";
        string filter = onlyActive ? "&$filter=statecode eq 0" : string.Empty;

        var all = new List<Dictionary<string, object?>>();
        int attempt = 0;
        while (true)
        {
            attempt++;
            bool restartRequested = false;
            string selectClause = string.Join(',', selectFields);
            string? url = baseUrl + "?$select=" + selectClause + filter;
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
        // Normalize null / empty
        if (a is string sa && string.IsNullOrWhiteSpace(sa)) a = null;
        if (b is string sb && string.IsNullOrWhiteSpace(sb)) b = null;
        if (a == null && b == null) return true;
        if (a == null || b == null) return false;

        // Guid comparison
        if (Guid.TryParse(a.ToString(), out var ga) && Guid.TryParse(b.ToString(), out var gb)) return ga.Equals(gb);

        // String normalization: trim and collapse internal whitespace; case-insensitive
        static string NormalizeString(string s)
        {
            // Trim and collapse multiple spaces to single space
            var trimmed = s.Trim();
            if (trimmed.Length == 0) return string.Empty;
            var parts = trimmed.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            return string.Join(' ', parts).ToLowerInvariant();
        }

        if (a is string || b is string)
        {
            var na = NormalizeString(a.ToString() ?? string.Empty);
            var nb = NormalizeString(b.ToString() ?? string.Empty);
            return na == nb;
        }

        // Fallback to case-insensitive string comparison of ToString values
        return string.Equals(a.ToString(), b.ToString(), StringComparison.OrdinalIgnoreCase);
    }

    private static byte[] BuildExcel(DiffReport report, string entityLogicalName, string sourceEnvUrl, string targetEnvUrl, DateTime comparisonTimestampUtc)
    {
        const int ExcelMaxCellLength = 32767;
        string Truncate(string? v)
        {
            if (string.IsNullOrEmpty(v)) return v ?? string.Empty;
            if (v.Length <= ExcelMaxCellLength) return v;
            return v.Substring(0, ExcelMaxCellLength - 15) + "...(truncated)";
        }

        // Helper to construct a direct model-driven app record URL (generic main.aspx link)
        string BuildRecordUrl(string envUrl, string etn, Guid id)
        {
            if (string.IsNullOrWhiteSpace(envUrl) || string.IsNullOrWhiteSpace(etn) || id == Guid.Empty) return string.Empty;
            // Using pagetype=entityrecord ensures navigation to record form; GUID without braces works.
            return envUrl.TrimEnd('/') + $"/main.aspx?etn={etn}&id={id}&pagetype=entityrecord";
        }

        using var wb = new XLWorkbook();
        // Summary sheet: single header row (requested columns) and single value row
        var summary = wb.Worksheets.Add("Summary");
        summary.Cell(1, 1).Value = "Source Env";
        summary.Cell(1, 2).Value = "Target Env";
        summary.Cell(1, 3).Value = "Comparission Date and Time"; // original spelling kept
        summary.Cell(1, 4).Value = "Entity";
        summary.Cell(1, 5).Value = "Compared Fields";
        summary.Cell(1, 6).Value = "Matching Records";
        summary.Cell(1, 7).Value = "Mismatches";
        summary.Cell(1, 8).Value = "Only In Source";
        summary.Cell(1, 9).Value = "Only In Target";
        summary.Cell(1, 10).Value = "Notes";
        summary.Cell(2, 1).Value = sourceEnvUrl;
        summary.Cell(2, 2).Value = targetEnvUrl;
        summary.Cell(2, 3).Value = comparisonTimestampUtc.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss");
        summary.Cell(2, 4).Value = entityLogicalName;
        summary.Cell(2, 5).Value = Truncate(string.Join(',', report.ComparedFields));
        summary.Cell(2, 6).Value = report.MatchingRecordIds.Count;
        summary.Cell(2, 7).Value = report.Mismatches.Count;
        summary.Cell(2, 8).Value = report.OnlyInSource.Count;
        summary.Cell(2, 9).Value = report.OnlyInTarget.Count;
        summary.Cell(2, 10).Value = Truncate(report.Notes);
        summary.Columns().AdjustToContents();

        var matchWs = wb.Worksheets.Add("MatchingRecords");
        matchWs.Cell(1, 1).Value = "Entity";
        matchWs.Cell(1, 2).Value = "RecordId";
        matchWs.Cell(1, 3).Value = "SourceRecordUrl";
        matchWs.Cell(1, 4).Value = "TargetRecordUrl";
        for (int i = 0; i < report.MatchingRecordIds.Count; i++)
        {
            var id = report.MatchingRecordIds[i];
            matchWs.Cell(i + 2, 1).Value = entityLogicalName;
            matchWs.Cell(i + 2, 2).Value = id.ToString();
            matchWs.Cell(i + 2, 3).Value = BuildRecordUrl(sourceEnvUrl, entityLogicalName, id);
            matchWs.Cell(i + 2, 4).Value = BuildRecordUrl(targetEnvUrl, entityLogicalName, id);
        }
        matchWs.Columns().AdjustToContents();

        var mismatchWs = wb.Worksheets.Add("Mismatches");
        mismatchWs.Cell(1, 1).Value = "Entity";
        mismatchWs.Cell(1, 2).Value = "RecordId";
        mismatchWs.Cell(1, 3).Value = "SourceRecordUrl";
        mismatchWs.Cell(1, 4).Value = "TargetRecordUrl";
        mismatchWs.Cell(1, 5).Value = "FieldNames";
        mismatchWs.Cell(1, 6).Value = "SourceValues";
        mismatchWs.Cell(1, 7).Value = "TargetValues";
        int r = 2;
        foreach (var grp in report.Mismatches.GroupBy(m => m.RecordId))
        {
            var recordId = grp.Key;
            var fieldNames = string.Join(',', grp.Select(g => g.FieldName));
            var sourceVals = string.Join(" | ", grp.Select(g => Truncate(g.SourceValue?.ToString())));
            var targetVals = string.Join(" | ", grp.Select(g => Truncate(g.TargetValue?.ToString())));
            mismatchWs.Cell(r, 1).Value = entityLogicalName;
            mismatchWs.Cell(r, 2).Value = recordId.ToString();
            mismatchWs.Cell(r, 3).Value = BuildRecordUrl(sourceEnvUrl, entityLogicalName, recordId);
            mismatchWs.Cell(r, 4).Value = BuildRecordUrl(targetEnvUrl, entityLogicalName, recordId);
            mismatchWs.Cell(r, 5).Value = Truncate(fieldNames);
            mismatchWs.Cell(r, 6).Value = Truncate(sourceVals);
            mismatchWs.Cell(r, 7).Value = Truncate(targetVals);
            r++;
        }
        mismatchWs.Columns().AdjustToContents();

        // Removed separate MultiFieldMismatches worksheet per request; individual field mismatches already captured above.

        // Write OnlyInSource / OnlyInTarget sheets with just RecordId column as requested
        void WriteRecordIdOnly(string sheetName, List<Dictionary<string, object?>> records, bool isSource)
        {
            var ws = wb.Worksheets.Add(sheetName);
            ws.Cell(1, 1).Value = "Entity";
            ws.Cell(1, 2).Value = "RecordId";
            ws.Cell(1, 3).Value = isSource ? "SourceRecordUrl" : "TargetRecordUrl";
            int row = 2;
            string pk = entityLogicalName + "id";
            var distinctIds = new HashSet<Guid>();
            foreach (var rec in records)
            {
                if (rec.TryGetValue(pk, out var idVal) && Guid.TryParse(idVal?.ToString(), out var gid))
                {
                    if (!distinctIds.Add(gid)) continue; // skip duplicates
                    ws.Cell(row, 1).Value = entityLogicalName;
                    ws.Cell(row, 2).Value = gid.ToString();
                    ws.Cell(row, 3).Value = BuildRecordUrl(isSource ? sourceEnvUrl : targetEnvUrl, entityLogicalName, gid);
                    row++;
                }
            }
            ws.Columns().AdjustToContents();
        }
        WriteRecordIdOnly("OnlyInSource", report.OnlyInSource, true);
        WriteRecordIdOnly("OnlyInTarget", report.OnlyInTarget, false);

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
            childMismatchWs.Cell(1, 3).Value = "SourceChildUrl";
            childMismatchWs.Cell(1, 4).Value = "TargetChildUrl";
            childMismatchWs.Cell(1, 5).Value = "FieldName";
            childMismatchWs.Cell(1, 6).Value = "SourceValue";
            childMismatchWs.Cell(1, 7).Value = "TargetValue";
            int cr = 2;
            foreach (var mm in child.Mismatches)
            {
                childMismatchWs.Cell(cr, 1).Value = mm.ParentId.ToString();
                childMismatchWs.Cell(cr, 2).Value = mm.ChildId.ToString();
                childMismatchWs.Cell(cr, 3).Value = BuildRecordUrl(sourceEnvUrl, child.ChildEntityLogicalName, mm.ChildId);
                childMismatchWs.Cell(cr, 4).Value = BuildRecordUrl(targetEnvUrl, child.ChildEntityLogicalName, mm.ChildId);
                childMismatchWs.Cell(cr, 5).Value = mm.FieldName;
                childMismatchWs.Cell(cr, 6).Value = Truncate(mm.SourceValue?.ToString());
                childMismatchWs.Cell(cr, 7).Value = Truncate(mm.TargetValue?.ToString());
                cr++;
            }
            childMismatchWs.Columns().AdjustToContents();

            void WriteChildRecordList(string sheetBase, Dictionary<Guid, List<Dictionary<string, object?>>> dict, bool isSource)
            {
                var ws = wb.Worksheets.Add($"Child_{child.ChildEntityLogicalName}_{sheetBase}".Replace('-', '_'));
                ws.Cell(1, 1).Value = "ParentId";
                ws.Cell(1, 2).Value = "ChildId";
                ws.Cell(1, 3).Value = isSource ? "SourceChildUrl" : "TargetChildUrl";
                for (int i = 0; i < child.ComparedChildFields.Count; i++) ws.Cell(1, i + 4).Value = child.ComparedChildFields[i];
                int row = 2;
                foreach (var parentPair in dict)
                {
                    foreach (var rec in parentPair.Value)
                    {
                        rec.TryGetValue(child.ChildEntityLogicalName + "id", out var cid);
                        Guid gid = Guid.TryParse(cid?.ToString(), out var tmp) ? tmp : Guid.Empty;
                        ws.Cell(row, 1).Value = parentPair.Key.ToString();
                        var cidCell = ws.Cell(row, 2);
                        cidCell.Value = gid == Guid.Empty ? cid?.ToString() : gid.ToString();
                        if (gid != Guid.Empty)
                        {
                            // Link cell shows GUID hyperlinked
                            var url = BuildRecordUrl(isSource ? sourceEnvUrl : targetEnvUrl, child.ChildEntityLogicalName, gid);
                            var linkCell = ws.Cell(row, 3);
                            linkCell.Value = url;
                            linkCell.SetHyperlink(new XLHyperlink(url));
                        }
                        for (int i = 0; i < child.ComparedChildFields.Count; i++)
                        {
                            rec.TryGetValue(child.ComparedChildFields[i], out var val);
                            ws.Cell(row, i + 4).Value = Truncate(val?.ToString());
                        }
                        row++;
                    }
                }
                ws.Columns().AdjustToContents();
            }
            if (child.OnlyInSourceByParent.Count > 0) WriteChildRecordList("OnlyInSource", child.OnlyInSourceByParent, true);
            if (child.OnlyInTargetByParent.Count > 0) WriteChildRecordList("OnlyInTarget", child.OnlyInTargetByParent, false);
        }

        using var ms = new System.IO.MemoryStream();
        wb.SaveAs(ms);
        return ms.ToArray();
    }

    private static string ExtractEnvShortName(string envUrl)
    {
        try
        {
            var uri = new Uri(envUrl);
            var host = uri.Host; // e.g. mash.crm.dynamics.com or mashppe.crm.dynamics.com
            var firstDot = host.IndexOf('.', StringComparison.OrdinalIgnoreCase);
            if (firstDot > 0) return host.Substring(0, firstDot);
            return host; // fallback
        }
        catch { return "env"; }
    }

    private static string GenerateBlobName(string sourceEnvUrl, string targetEnvUrl, string entityLogicalName, bool isCombined = false)
    {
        // New naming convention requested:
        // D365Config_Comparision_<SourceShort>_Vs_<TargetShort>_<Timestamp>.xlsx
        // Applies to both single-entity and multi-entity outputs. Entity name no longer embedded.
        var srcShort = ExtractEnvShortName(sourceEnvUrl);
        var tgtShort = ExtractEnvShortName(targetEnvUrl);
        var ts = DateTime.UtcNow.ToString("yyyyMMddHHmmss");
        return $"D365Config_Comparision_{srcShort}_Vs_{tgtShort}_{ts}.xlsx";
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
            string? url = baseUrl + "?$select=" + selectClause + filter;
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
