using System;
using System.Buffers;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Logging;
using ClosedXML.Excel;
using Azure.Storage.Blobs;
using Azure.Core;
using Azure.Identity;
using Microsoft.Azure.Functions.Worker.Extensions.Timer;

namespace D365ComparissionTool;

internal class FlowComparisonRequestPayload
{
    public string? Mode { get; set; } // all|single|list
    public string? Env { get; set; }
    public string? Env2 { get; set; }
    public string? Token1 { get; set; } // bearer token for primary env (now provided in body)
    public string? Token2 { get; set; }
    public string? MatchBy { get; set; } // name|id (default name)
    public bool IncludeDiffDetails { get; set; } = true;
    public bool StoreSnapshots { get; set; } = false; // placeholder (blob not implemented yet)
    public string? FlowId { get; set; }
    public string? FlowName { get; set; }
    public List<string>? Flows { get; set; }
    // Storage configuration (optional) similar to CompareDataverse
    public string? StorageAccountUrl { get; set; } // e.g. https://account.blob.core.windows.net
    public string? OutputContainerName { get; set; }
    public string? OutputFileName { get; set; } // optional override
    public string? StorageConnectionString { get; set; }
    public string? StorageContainerSasUrl { get; set; } // full https://.../container?sv=...
}

internal record FlowSnapshot(string Name, Guid FlowId, string Environment, string CanonicalJson, string Hash, string? Error);

internal record FlowComparisonDiff(List<string> Added, List<string> Removed, List<(string Path, string OldValue, string NewValue)> Changed);

internal record FlowActionDifference(string ActionName, string Status, List<(string Path, string OldValue, string NewValue)> ChangedProperties); // Status: added|removed|changed

internal record FlowComparisonResult(string Name,
    FlowSnapshot? Env1,
    FlowSnapshot? Env2,
    bool Identical,
    FlowComparisonDiff? Diff,
    List<FlowActionDifference>? ActionDifferences,
    string Status); // identical|different|missing_in_env2|error

internal record FlowComparisonResponse(List<FlowSnapshot> Env1Flows,
    List<FlowSnapshot> Env2Flows,
    List<FlowComparisonResult> Comparisons,
    List<string> IdenticalFlows,
    List<string> NonIdenticalFlows,
    int MissingInEnv2Count,
    int ErrorCount,
    string Notes,
    string? ExcelBlobUrl,
    bool Env1AuthFailed,
    bool Env2AuthFailed);

public class FlowComparisonFunction
{
    private readonly ILogger _logger;
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true
    };

    public FlowComparisonFunction(ILogger<FlowComparisonFunction> logger) => _logger = logger;

    // Weekly timer trigger to compare flows between source and multiple target environments
    // Schedule: Friday 09:00 UTC (same as configuration comparison)
    [Function("CompareFlowsWeekly")]
    public async Task RunFlowsWeekly([TimerTrigger("0 0 9 * * Fri")] TimerInfo timerInfo, FunctionContext context)
    {
        var correlationId = Guid.NewGuid();
        _logger.LogInformation("[CompareFlowsWeekly] Timer fired (Fri 09:00 UTC) CorrelationId={CorrelationId} LastRun={LastRun} NextRun={NextRun}", correlationId, timerInfo?.ScheduleStatus?.Last, timerInfo?.ScheduleStatus?.Next);
        try
        {
            // Source environment URL (full) and host
            var srcEnvUrl = Environment.GetEnvironmentVariable("FLOW_SOURCE_ENV_URL")
                            ?? Environment.GetEnvironmentVariable("SOURCE_ENV_URL")
                            ?? string.Empty;
            if (string.IsNullOrWhiteSpace(srcEnvUrl))
            {
                _logger.LogWarning("[{CorrelationId}] FLOW_SOURCE_ENV_URL/SOURCE_ENV_URL missing; aborting flow weekly run", correlationId);
                return;
            }
            var srcHost = NormalizeEnvHost(srcEnvUrl);

            // Determine target environment list from env var FLOW_TARGET_ENV_LIST (JSON array or comma-separated)
            var targetsRaw = Environment.GetEnvironmentVariable("FLOW_TARGET_ENV_LIST");
            var targetHosts = new List<string>();
            if (!string.IsNullOrWhiteSpace(targetsRaw))
            {
                bool parsed = false;
                try
                {
                    var arr = JsonSerializer.Deserialize<List<string>>(targetsRaw);
                    if (arr != null && arr.Count > 0)
                    {
                        targetHosts.AddRange(arr.Where(a => !string.IsNullOrWhiteSpace(a)));
                        parsed = true;
                    }
                }
                catch { }
                if (!parsed)
                {
                    targetHosts.AddRange(targetsRaw.Split(new[] { ',', ';', ' ' }, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries));
                }
            }
            // No hardcoded fallback targets; require FLOW_TARGET_ENV_LIST or handle single target via SOURCE/TARGET envs.
            targetHosts = targetHosts.Select(h => NormalizeEnvHost(h)).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
            _logger.LogInformation("[{CorrelationId}] Flow weekly comparison SourceHost={SourceHost} TargetCount={Count} Targets={Targets}", correlationId, srcHost, targetHosts.Count, string.Join(',', targetHosts));

            // Token acquisition via DefaultAzureCredential per environment host
            var credential = new DefaultAzureCredential();
            string sourceToken = await AcquireTokenForEnvAsync(srcHost, credential);
            _logger.LogInformation("[{CorrelationId}] Acquired source token length={Len}", correlationId, sourceToken.Length);

            // Shared ignore keys (same logic as HTTP function)
            var ignoreKeysConfig = Environment.GetEnvironmentVariable("NORMALIZE_IGNORE_KEYS") ?? string.Empty;
            var ignoreKeys = ignoreKeysConfig.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                .Select(k => k.Trim()).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
            if (ignoreKeys.Count == 0)
            {
                ignoreKeys.AddRange(new[] { "connectionReferences", "runtimeConfiguration", "lastModified", "createdTime", "modifiedTime", "etag", "trackedProperties" });
            }

            foreach (var tgtHost in targetHosts)
            {
                try
                {
                    string targetToken = await AcquireTokenForEnvAsync(tgtHost, credential);
                    _logger.LogInformation("[{CorrelationId}] Processing target host={TargetHost} tokenLength={Len}", correlationId, tgtHost, targetToken.Length);

                    // Fetch flows for source and target
                    var env1Flows = await FetchAndPrepareAllFlows(srcHost, sourceToken, ignoreKeys, correlationId);
                    var env2Flows = await FetchAndPrepareAllFlows(tgtHost, targetToken, ignoreKeys, correlationId);
                    var comparisons = new List<FlowComparisonResult>();

                    // Build lookup maps for target by name
                    var env2LookupByName = env2Flows.GroupBy(f => f.Name, StringComparer.OrdinalIgnoreCase)
                        .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);

                    foreach (var f1 in env1Flows)
                    {
                        env2LookupByName.TryGetValue(f1.Name, out var f2);
                        if (f2 == null)
                        {
                            comparisons.Add(new FlowComparisonResult(f1.Name, f1, null, false, null, null, "missing_in_env2"));
                            continue;
                        }
                        if (!string.IsNullOrEmpty(f1.Error) || !string.IsNullOrEmpty(f2.Error))
                        {
                            comparisons.Add(new FlowComparisonResult(f1.Name, f1, f2, false, null, null, "error"));
                            continue;
                        }
                        bool identical = string.Equals(f1.Hash, f2.Hash, StringComparison.OrdinalIgnoreCase);
                        FlowComparisonDiff? diff = null;
                        List<FlowActionDifference>? actionDiffs = null;
                        if (!identical)
                        {
                            diff = ComputeDiff(f1.CanonicalJson, f2.CanonicalJson);
                            actionDiffs = ComputeActionDifferences(f1.CanonicalJson, f2.CanonicalJson);
                        }
                        comparisons.Add(new FlowComparisonResult(f1.Name, f1, f2, identical, diff, actionDiffs, identical ? "identical" : "different"));
                    }

                    // Build Excel and upload
                    string? excelBlobUrl = null;
                    try
                    {
                        var wbBytes = BuildFlowComparisonExcel(env1Flows, env2Flows, comparisons, srcHost, tgtHost, DateTime.UtcNow);
                        string blobName = GenerateFlowBlobName(srcHost, tgtHost, null);
                        string sas = Environment.GetEnvironmentVariable("COMPARISON_STORAGE_CONTAINER_SAS_URL") ?? string.Empty;
                        if (!string.IsNullOrWhiteSpace(sas))
                        {
                            excelBlobUrl = await UploadExcelWithFlexibleAuthAsync(wbBytes, blobName, containerSasUrl: sas);
                        }
                        else
                        {
                            _logger.LogWarning("[{CorrelationId}] SAS URL missing; Excel not uploaded for target={TargetHost}", correlationId, tgtHost);
                        }
                        if (!string.IsNullOrWhiteSpace(excelBlobUrl))
                        {
                            _logger.LogInformation("[{CorrelationId}] Flow weekly Excel uploaded Target={TargetHost} BlobUrl={BlobUrl}", correlationId, tgtHost, excelBlobUrl);
                        }
                    }
                    catch (Exception exUp)
                    {
                        _logger.LogError(exUp, "[{CorrelationId}] Excel generation/upload failed for target={TargetHost}", correlationId, tgtHost);
                    }
                }
                catch (Exception exTarget)
                {
                    _logger.LogError(exTarget, "[{CorrelationId}] Flow weekly comparison failed for target host={TargetHost}", correlationId, tgtHost);
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "[CompareFlowsWeekly] Unhandled error CorrelationId={CorrelationId}", correlationId);
        }
    }

    private static string NormalizeEnvHost(string env)
    {
        var e = env.Trim();
        if (e.StartsWith("https://", StringComparison.OrdinalIgnoreCase)) e = e[8..];
        else if (e.StartsWith("http://", StringComparison.OrdinalIgnoreCase)) e = e[7..];
        // Remove any trailing path beyond host if accidentally supplied
        int slash = e.IndexOf('/')
;        if (slash >= 0) e = e[..slash];
        return e.Trim();
    }

    private static async Task<string> AcquireTokenForEnvAsync(string envHost, DefaultAzureCredential credential)
    {
        // Scope uses full https://host/.default
        var scope = $"https://{envHost.TrimEnd('/')}/.default";
        var token = await credential.GetTokenAsync(new TokenRequestContext(new[] { scope }));
        return token.Token;
    }

    private async Task<List<FlowSnapshot>> FetchAndPrepareAllFlows(string env, string token, List<string> ignoreKeys, Guid correlationId)
    {
        var list = new List<FlowSnapshot>();
        _logger.LogInformation("[CompareFlows:{CorrelationId}] FetchAll start Env={Env}", correlationId, env);
        foreach (var raw in await FetchFlowsRaw(env, token, correlationId))
        {
            var snap = BuildSnapshot(raw, env, ignoreKeys);
            list.Add(snap);
        }
        _logger.LogInformation("[CompareFlows:{CorrelationId}] FetchAll end Env={Env} Count={Count}", correlationId, env, list.Count);
        return list;
    }

    private async Task<FlowSnapshot?> FetchSingleFlow(string env, string token, string identifier, List<string> ignoreKeys, Guid correlationId, string? matchBy = null)
    {
        var rawList = await FetchFlowsRaw(env, token, correlationId, identifier);
        var raw = rawList.FirstOrDefault();
        if (raw == null) return null;
        return BuildSnapshot(raw, env, ignoreKeys);
    }

    private record RawFlow(Guid FlowId, string Name, string? ClientDataJson, JsonDocument? ClientDataDoc);

    private async Task<List<RawFlow>> FetchFlowsRaw(string env, string token, Guid correlationId, string? singleIdentifier = null)
    {
        var client = new HttpClient();
        // Add required OData and User-Agent headers for Dataverse API
        if (!client.DefaultRequestHeaders.Contains("OData-MaxVersion"))
            client.DefaultRequestHeaders.Add("OData-MaxVersion", "4.0");
        if (!client.DefaultRequestHeaders.Contains("OData-Version"))
            client.DefaultRequestHeaders.Add("OData-Version", "4.0");
        if (!client.DefaultRequestHeaders.Contains("User-Agent"))
            client.DefaultRequestHeaders.Add("User-Agent", "D365ComparisonTool/1.0");
        client.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);
        client.DefaultRequestHeaders.Accept.ParseAdd("application/json");
        // Query workflows entity (cloud flows); category 5 = modern flows
        string baseUrl = $"https://{env.TrimEnd('/')}/api/data/v9.2/workflows?$select=workflowid,name,clientdata&$filter=category eq 5";
        _logger.LogInformation("[CompareFlows:{CorrelationId}] FetchFlowsRaw start Env={Env} SingleIdentifier={Identifier}", correlationId, env, singleIdentifier ?? "(all)");
        if (!string.IsNullOrWhiteSpace(singleIdentifier))
        {
            if (Guid.TryParse(singleIdentifier, out var gid))
                baseUrl = $"https://{env.TrimEnd('/')}/api/data/v9.2/workflows({gid})?$select=workflowid,name,clientdata";
            else
            {
                // override baseUrl without category filter to allow name filter (append existing category condition)
                baseUrl = $"https://{env.TrimEnd('/')}/api/data/v9.2/workflows?$select=workflowid,name,clientdata&$filter=category eq 5 and name eq '{singleIdentifier.Replace("'", "''")}'";
            }
        }
        var list = new List<RawFlow>();
        string? url = baseUrl;
        int page = 0;
        while (!string.IsNullOrEmpty(url))
        {
            page++;
            using var req = new HttpRequestMessage(HttpMethod.Get, url);
            using var resp = await client.SendAsync(req);
            if (!resp.IsSuccessStatusCode)
            {
                var contentSnippet = string.Empty;
                try
                {
                    contentSnippet = (await resp.Content.ReadAsStringAsync()) ?? string.Empty;
                    if (contentSnippet.Length > 500) contentSnippet = contentSnippet.Substring(0, 500) + "...";
                }
                catch { }
                if ((int)resp.StatusCode == 429 || (int)resp.StatusCode >= 500)
                {
                    _logger.LogWarning("[CompareFlows:{CorrelationId}] Transient status={Status} page={Page} backing off BodySnippet={Body}", correlationId, resp.StatusCode, page, contentSnippet);
                    await Task.Delay(Math.Min(5000, 250 * page));
                    continue;
                }
                _logger.LogError("[CompareFlows:{CorrelationId}] Non-success status={Status} page={Page} Url={Url} Body={Body}", correlationId, resp.StatusCode, page, url, contentSnippet);
                throw new HttpRequestException($"FetchFlowsRaw failed (status {(int)resp.StatusCode} {resp.StatusCode}) page={page} url={url} body={contentSnippet}");
            }
            using var stream = await resp.Content.ReadAsStreamAsync();
            using var doc = await JsonDocument.ParseAsync(stream);
            if (doc.RootElement.ValueKind == JsonValueKind.Object && singleIdentifier != null && doc.RootElement.TryGetProperty("flowid", out var fid))
            {
                var raw = ParseSingle(doc.RootElement);
                list.Add(raw);
                break;
            }
            if (doc.RootElement.TryGetProperty("value", out var arr))
            {
                foreach (var item in arr.EnumerateArray()) list.Add(ParseSingle(item));
            }
            url = doc.RootElement.TryGetProperty("@odata.nextLink", out var next) ? next.GetString() : null;
            _logger.LogDebug("[CompareFlows:{CorrelationId}] PageComplete Env={Env} Page={Page} AccumulatedCount={Count} HasNext={HasNext}", correlationId, env, page, list.Count, url != null);
        }
        _logger.LogInformation("[CompareFlows:{CorrelationId}] FetchFlowsRaw end Env={Env} TotalFlows={Count}", correlationId, env, list.Count);
        return list;

        RawFlow ParseSingle(JsonElement el)
        {
            Guid id = Guid.Empty; string name = string.Empty; string? clientdata = null;
            // workflowid is the correct primary key; retain backward compatibility if 'flowid' were ever present
            if (el.TryGetProperty("workflowid", out var wfId) && Guid.TryParse(wfId.GetString(), out var g1)) id = g1;
            else if (el.TryGetProperty("flowid", out var fid) && Guid.TryParse(fid.GetString(), out var g2)) id = g2;
            if (el.TryGetProperty("name", out var nm)) name = nm.GetString() ?? string.Empty;
            if (el.TryGetProperty("clientdata", out var cd))
            {
                if (cd.ValueKind == JsonValueKind.String) clientdata = cd.GetString();
                else clientdata = cd.GetRawText();
            }
            JsonDocument? clientDoc = null;
            if (!string.IsNullOrWhiteSpace(clientdata))
            {
                try { clientDoc = JsonDocument.Parse(clientdata); } catch { }
            }
            return new RawFlow(id, name, clientdata, clientDoc);
        }
    }

    private FlowSnapshot BuildSnapshot(RawFlow raw, string env, List<string> ignoreKeys)
    {
        if (raw.ClientDataDoc == null)
        {
            _logger.LogWarning("[CompareFlows] Snapshot build skipped (no clientdata) Env={Env} Name={Name} Id={Id}", env, raw.Name, raw.FlowId);
            return new FlowSnapshot(raw.Name, raw.FlowId, env, string.Empty, string.Empty, "clientdata missing or invalid");
        }
        try
        {
            var normalized = NormalizeFlowDefinition(raw.ClientDataDoc.RootElement, ignoreKeys, env);
            string canonical = CanonicalSerialize(normalized);
            string hash = ComputeSha256(canonical);
            _logger.LogDebug("[CompareFlows] Snapshot built Env={Env} Name={Name} Id={Id} Hash={Hash} CanonicalLength={Len}", env, raw.Name, raw.FlowId, hash, canonical.Length);
            return new FlowSnapshot(raw.Name, raw.FlowId, env, canonical, hash, null);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "[CompareFlows] Snapshot build error Env={Env} Name={Name} Id={Id}", env, raw.Name, raw.FlowId);
            return new FlowSnapshot(raw.Name, raw.FlowId, env, string.Empty, string.Empty, ex.Message);
        }
    }

    private static JsonElement NormalizeFlowDefinition(JsonElement root, List<string> ignoreKeys, string envHost)
    {
        using var doc = JsonDocument.Parse(root.GetRawText());
        var toRemove = new HashSet<string>(ignoreKeys, StringComparer.OrdinalIgnoreCase);
        // Recursively build normalized structure using mutable dictionary/list then serialize back
        object? Traverse(JsonElement el)
        {
            switch (el.ValueKind)
            {
                case JsonValueKind.Object:
                    var dict = new SortedDictionary<string, object?>(StringComparer.Ordinal); // lexicographic
                    foreach (var prop in el.EnumerateObject())
                    {
                        if (ShouldIgnore(prop.Name, toRemove)) continue;
                        var val = Traverse(prop.Value);
                        dict[prop.Name] = val;
                    }
                    return dict;
                case JsonValueKind.Array:
                    var arr = new List<object?>();
                    foreach (var item in el.EnumerateArray()) arr.Add(Traverse(item));
                    return arr;
                case JsonValueKind.String:
                    var s = el.GetString() ?? string.Empty;
                    s = MaskGuids(s);
                    s = MaskHostUrls(s, envHost);
                    return s;
                case JsonValueKind.Number:
                    if (el.TryGetInt64(out var l)) return l;
                    if (el.TryGetDouble(out var d)) return d;
                    return el.GetRawText();
                case JsonValueKind.True: return true;
                case JsonValueKind.False: return false;
                case JsonValueKind.Null: return null;
                default: return el.GetRawText();
            }
        }
        var normalizedObj = Traverse(doc.RootElement);
        // Serialize normalizedObj back into JsonElement
        string json = CanonicalSerialize(normalizedObj);
        using var finalDoc = JsonDocument.Parse(json);
        return finalDoc.RootElement.Clone();
    }

    private static bool ShouldIgnore(string name, HashSet<string> ignore)
    {
        if (ignore.Contains(name)) return true;
        // pattern based ignores
        if (Regex.IsMatch(name, ".*id$", RegexOptions.IgnoreCase)) return true;
        if (Regex.IsMatch(name, ".*Id$", RegexOptions.IgnoreCase)) return true; // duplicate but explicit
        if (name.StartsWith("connection", StringComparison.OrdinalIgnoreCase)) return true;
        return false;
    }

    private static string MaskGuids(string input)
    {
        return Regex.Replace(input, "[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}", "<GUID>");
    }

    private static string MaskHostUrls(string input, string host)
    {
        if (string.IsNullOrWhiteSpace(host)) return input;
        return input.Replace(host, "<ENV_HOST>", StringComparison.OrdinalIgnoreCase);
    }

    private static string CanonicalSerialize(object? obj)
    {
        var buffer = new ArrayBufferWriter<byte>();
        // Utf8JsonWriter implements IDisposable; keep using for writer only
        using (var writer = new Utf8JsonWriter(buffer, new JsonWriterOptions { Indented = false }))
        {
            WriteCanonical(writer, obj);
        }
        return Encoding.UTF8.GetString(buffer.WrittenSpan);
    }

    private static void WriteCanonical(Utf8JsonWriter writer, object? obj)
    {
        switch (obj)
        {
            case null:
                writer.WriteNullValue(); break;
            case string s:
                writer.WriteStringValue(s); break;
            case bool b:
                writer.WriteBooleanValue(b); break;
            case int i:
                writer.WriteNumberValue(i); break;
            case long l:
                writer.WriteNumberValue(l); break;
            case double d:
                writer.WriteNumberValue(d); break;
            case SortedDictionary<string, object?> dict:
                writer.WriteStartObject();
                foreach (var kv in dict) { writer.WritePropertyName(kv.Key); WriteCanonical(writer, kv.Value); }
                writer.WriteEndObject();
                break;
            case Dictionary<string, object?> d2:
                // ensure ordering
                writer.WriteStartObject();
                foreach (var kv in d2.OrderBy(k => k.Key, StringComparer.Ordinal)) { writer.WritePropertyName(kv.Key); WriteCanonical(writer, kv.Value); }
                writer.WriteEndObject();
                break;
            case IEnumerable<object?> list:
                writer.WriteStartArray();
                foreach (var item in list) WriteCanonical(writer, item);
                writer.WriteEndArray();
                break;
            default:
                if (obj is JsonElement je)
                {
                    je.WriteTo(writer);
                }
                else
                {
                    writer.WriteStringValue(obj.ToString());
                }
                break;
        }
    }

    private static string ComputeSha256(string canonical)
    {
        using var sha = SHA256.Create();
        var bytes = sha.ComputeHash(Encoding.UTF8.GetBytes(canonical));
        return BitConverter.ToString(bytes).Replace("-", string.Empty).ToLowerInvariant();
    }

    // Excel helpers (outside comparison loop)
    private static byte[] BuildFlowComparisonExcel(List<FlowSnapshot> env1Flows, List<FlowSnapshot> env2Flows, List<FlowComparisonResult> comparisons, string env1Host, string? env2Host, DateTime comparisonTimestampUtc)
    {
        using var wb = new XLWorkbook();
        var summary = wb.Worksheets.Add("Summary");
        summary.Cell(1,1).Value = "Env1";
        summary.Cell(1,2).Value = "Env2";
        summary.Cell(1,3).Value = "Comparison Date and Time";
        summary.Cell(1,4).Value = "Env1 Flow Count";
        summary.Cell(1,5).Value = "Env2 Flow Count";
        summary.Cell(1,6).Value = "Identical";
        summary.Cell(1,7).Value = "Different";
        summary.Cell(1,8).Value = "Missing In Env2";
        summary.Cell(1,9).Value = "Errors";
        summary.Cell(2,1).Value = env1Host;
        summary.Cell(2,2).Value = env2Host ?? "(single-env)";
        summary.Cell(2,3).Value = comparisonTimestampUtc.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss");
        summary.Cell(2,4).Value = env1Flows.Count;
        summary.Cell(2,5).Value = env2Flows.Count;
        summary.Cell(2,6).Value = comparisons.Count(c => c.Identical);
        summary.Cell(2,7).Value = comparisons.Count(c => c.Status == "different");
        summary.Cell(2,8).Value = comparisons.Count(c => c.Status == "missing_in_env2");
        summary.Cell(2,9).Value = comparisons.Count(c => c.Status == "error");
        summary.Columns().AdjustToContents();

        var overview = wb.Worksheets.Add("FlowOverview");
        overview.Cell(1,1).Value = "Name";
        overview.Cell(1,2).Value = "Status";
        overview.Cell(1,3).Value = "Identical";
        overview.Cell(1,4).Value = "Env1 FlowId";
        overview.Cell(1,5).Value = "Env2 FlowId";
        overview.Cell(1,6).Value = "Env1 Hash";
        overview.Cell(1,7).Value = "Env2 Hash";
        overview.Cell(1,8).Value = "AddedPaths";
        overview.Cell(1,9).Value = "RemovedPaths";
        overview.Cell(1,10).Value = "ChangedPaths";
        overview.Cell(1,11).Value = "ActionDiffs";
        overview.Cell(1,12).Value = "Env1Error";
        overview.Cell(1,13).Value = "Env2Error";
        int or = 2;
        foreach (var cmp in comparisons)
        {
            overview.Cell(or,1).Value = cmp.Name;
            overview.Cell(or,2).Value = cmp.Status;
            overview.Cell(or,3).Value = cmp.Identical ? 1 : 0;
            overview.Cell(or,4).Value = cmp.Env1?.FlowId.ToString() ?? string.Empty;
            overview.Cell(or,5).Value = cmp.Env2?.FlowId.ToString() ?? string.Empty;
            overview.Cell(or,6).Value = cmp.Env1?.Hash ?? string.Empty;
            overview.Cell(or,7).Value = cmp.Env2?.Hash ?? string.Empty;
            overview.Cell(or,8).Value = cmp.Diff?.Added.Count ?? 0;
            overview.Cell(or,9).Value = cmp.Diff?.Removed.Count ?? 0;
            overview.Cell(or,10).Value = cmp.Diff?.Changed.Count ?? 0;
            overview.Cell(or,11).Value = cmp.ActionDifferences?.Count ?? 0;
            overview.Cell(or,12).Value = cmp.Env1?.Error ?? string.Empty;
            overview.Cell(or,13).Value = cmp.Env2?.Error ?? string.Empty;
            or++;
        }
        overview.Columns().AdjustToContents();

        var diffWs = wb.Worksheets.Add("Differences");
        diffWs.Cell(1,1).Value = "FlowName";
        diffWs.Cell(1,2).Value = "PathType";
        diffWs.Cell(1,3).Value = "Path";
        diffWs.Cell(1,4).Value = "OldValue";
        diffWs.Cell(1,5).Value = "NewValue";
        int dr = 2;
        foreach (var cmp in comparisons.Where(c => c.Status == "different" && c.Diff != null))
        {
            foreach (var a in cmp.Diff!.Added) { diffWs.Cell(dr,1).Value = cmp.Name; diffWs.Cell(dr,2).Value = "Added"; diffWs.Cell(dr,3).Value = a; diffWs.Cell(dr,4).Value = string.Empty; diffWs.Cell(dr,5).Value = "(added)"; dr++; }
            foreach (var r in cmp.Diff!.Removed) { diffWs.Cell(dr,1).Value = cmp.Name; diffWs.Cell(dr,2).Value = "Removed"; diffWs.Cell(dr,3).Value = r; diffWs.Cell(dr,4).Value = "(removed)"; diffWs.Cell(dr,5).Value = string.Empty; dr++; }
            foreach (var ch in cmp.Diff!.Changed) { diffWs.Cell(dr,1).Value = cmp.Name; diffWs.Cell(dr,2).Value = "Changed"; diffWs.Cell(dr,3).Value = ch.Path; diffWs.Cell(dr,4).Value = TrimForExcel(ch.OldValue); diffWs.Cell(dr,5).Value = TrimForExcel(ch.NewValue); dr++; }
        }
        diffWs.Columns().AdjustToContents();

        var actionWs = wb.Worksheets.Add("ActionDifferences");
        actionWs.Cell(1,1).Value = "FlowName";
        actionWs.Cell(1,2).Value = "ActionName";
        actionWs.Cell(1,3).Value = "Status";
        actionWs.Cell(1,4).Value = "ChangedPropertyCount";
        actionWs.Cell(1,5).Value = "ChangedProperties (Path=Old=>New)";
        int ar = 2;
        foreach (var cmp in comparisons.Where(c => c.ActionDifferences != null && c.ActionDifferences.Count > 0))
        {
            foreach (var ad in cmp.ActionDifferences!)
            {
                actionWs.Cell(ar,1).Value = cmp.Name;
                actionWs.Cell(ar,2).Value = ad.ActionName;
                actionWs.Cell(ar,3).Value = ad.Status;
                actionWs.Cell(ar,4).Value = ad.ChangedProperties.Count;
                var combined = string.Join(" | ", ad.ChangedProperties.Select(cp => cp.Path + "=" + TrimForExcel(cp.OldValue) + "=>" + TrimForExcel(cp.NewValue)));
                actionWs.Cell(ar,5).Value = TrimForExcel(combined);
                ar++;
            }
        }
        actionWs.Columns().AdjustToContents();

        using var ms = new MemoryStream();
        wb.SaveAs(ms);
        return ms.ToArray();

        static string TrimForExcel(string? v)
        {
            if (string.IsNullOrEmpty(v)) return string.Empty;
            const int ExcelMaxCellLength = 32767;
            return v.Length <= ExcelMaxCellLength ? v : v.Substring(0, ExcelMaxCellLength - 15) + "...(truncated)";
        }
    }

    private static string GenerateFlowBlobName(string env1Host, string? env2Host, string? overrideName)
    {
        if (!string.IsNullOrWhiteSpace(overrideName))
        {
            var safe = new string(overrideName.Where(ch => char.IsLetterOrDigit(ch) || ch=='-' || ch=='_' || ch=='.').ToArray());
            if (!safe.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)) safe += ".xlsx";
            return safe;
        }
        // New naming convention requested:
        // D365Flow_Comparision_<SourceShort>_Vs_<TargetShort>_<Timestamp>.xlsx
        // If single env (no env2) -> D365Flow_Comparision_<SourceShort>_<Timestamp>.xlsx
        var ts = DateTime.UtcNow.ToString("yyyyMMddHHmmss");
        string srcShort = ExtractShortHost(env1Host);
        if (string.IsNullOrWhiteSpace(env2Host))
        {
            return $"D365Flow_Comparision_{srcShort}_{ts}.xlsx";
        }
        string tgtShort = ExtractShortHost(env2Host);
        return $"D365Flow_Comparision_{srcShort}_Vs_{tgtShort}_{ts}.xlsx";
    }

    private static string ExtractShortHost(string host)
    {
        if (string.IsNullOrWhiteSpace(host)) return "env";
        var h = host.Trim();
        // remove protocol if accidentally included
        if (h.StartsWith("https://", StringComparison.OrdinalIgnoreCase)) h = h[8..];
        else if (h.StartsWith("http://", StringComparison.OrdinalIgnoreCase)) h = h[7..];
        int dot = h.IndexOf('.');
        if (dot > 0) return h.Substring(0, dot);
        return h;
    }

    private static async Task<string> UploadExcelWithFlexibleAuthAsync(byte[] bytes, string blobName, string? accountUrl = null, string? containerName = null, TokenCredential? credential = null, string? connectionString = null, string? containerSasUrl = null)
    {
        BlobContainerClient containerClient;
        if (!string.IsNullOrWhiteSpace(containerSasUrl))
        {
            var raw = containerSasUrl.Trim();
            if (!Uri.TryCreate(raw, UriKind.Absolute, out var sasUri)) throw new ArgumentException("StorageContainerSasUrl invalid absolute URI");
            if (string.IsNullOrEmpty(sasUri.Query)) throw new ArgumentException("StorageContainerSasUrl missing SAS query params");
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
        using var ms = new MemoryStream(bytes);
        var blob = containerClient.GetBlobClient(blobName);
        await blob.UploadAsync(ms, overwrite: true);
        return blob.Uri.ToString();
    }

    private static FlowComparisonDiff ComputeDiff(string jsonA, string jsonB)
    {
        using var docA = JsonDocument.Parse(jsonA);
        using var docB = JsonDocument.Parse(jsonB);
        var added = new List<string>();
        var removed = new List<string>();
        var changed = new List<(string Path, string OldValue, string NewValue)>();
        CompareElements(docA.RootElement, docB.RootElement, "$", added, removed, changed);
        return new FlowComparisonDiff(added, removed, changed);
    }

    private static List<FlowActionDifference>? ComputeActionDifferences(string jsonA, string jsonB)
    {
        try
        {
            using var docA = JsonDocument.Parse(jsonA);
            using var docB = JsonDocument.Parse(jsonB);
            var actionsA = ExtractActionsMap(docA.RootElement);
            var actionsB = ExtractActionsMap(docB.RootElement);
            if (actionsA == null || actionsB == null) return null; // actions not present
            var diffs = new List<FlowActionDifference>();
            // Removed actions
            foreach (var name in actionsA.Keys.Except(actionsB.Keys))
            {
                diffs.Add(new FlowActionDifference(name, "removed", new List<(string, string, string)>()));
            }
            // Added actions
            foreach (var name in actionsB.Keys.Except(actionsA.Keys))
            {
                diffs.Add(new FlowActionDifference(name, "added", new List<(string, string, string)>()));
            }
            // Changed actions
            foreach (var name in actionsA.Keys.Intersect(actionsB.Keys))
            {
                var aEl = actionsA[name];
                var bEl = actionsB[name];
                var aText = aEl.GetRawText();
                var bText = bEl.GetRawText();
                if (aText != bText)
                {
                    var changedProps = new List<(string Path, string OldValue, string NewValue)>();
                    // Capture added / removed so we can decide if the difference is ONLY a runAfter reference change (dependency rename)
                    var added = new List<string>();
                    var removed = new List<string>();
                    CompareElements(aEl, bEl, "$..actions." + name, added, removed, changedProps);
                    // Surface all configurational additions/removals (not only runAfter) as placeholder change entries
                    // Avoid duplicating paths already captured in changedProps for value changes.
                    var existingPaths = new HashSet<string>(changedProps.Select(cp => cp.Path));
                    foreach (var p in added)
                    {
                        if (!existingPaths.Contains(p))
                            changedProps.Add((p, "(missing)", "(added)"));
                    }
                    foreach (var p in removed)
                    {
                        if (!existingPaths.Contains(p))
                            changedProps.Add((p, "(removed)", "(missing)"));
                    }
                    diffs.Add(new FlowActionDifference(name, "changed", changedProps));
                }
            }
            return diffs;
        }
        catch
        {
            return null; // fail silently, higher-level diff still available
        }
    }

    private static Dictionary<string, JsonElement>? ExtractActionsMap(JsonElement root)
    {
        // Try common paths: properties.definition.actions OR definition.actions
        JsonElement definition;
        if (root.ValueKind == JsonValueKind.Object)
        {
            if (root.TryGetProperty("properties", out var props) && props.ValueKind == JsonValueKind.Object && props.TryGetProperty("definition", out var def) && def.ValueKind == JsonValueKind.Object)
            {
                definition = def;
                if (definition.TryGetProperty("actions", out var actions) && actions.ValueKind == JsonValueKind.Object)
                    return actions.EnumerateObject().ToDictionary(p => p.Name, p => p.Value, StringComparer.OrdinalIgnoreCase);
            }
            if (root.TryGetProperty("definition", out var def2) && def2.ValueKind == JsonValueKind.Object)
            {
                definition = def2;
                if (definition.TryGetProperty("actions", out var actions2) && actions2.ValueKind == JsonValueKind.Object)
                    return actions2.EnumerateObject().ToDictionary(p => p.Name, p => p.Value, StringComparer.OrdinalIgnoreCase);
            }
        }
        return null;
    }

    private static void CompareElements(JsonElement a, JsonElement b, string path, List<string> added, List<string> removed, List<(string, string, string)> changed)
    {
        if (a.ValueKind != b.ValueKind)
        {
            changed.Add((path, a.GetRawText(), b.GetRawText()));
            return;
        }
        switch (a.ValueKind)
        {
            case JsonValueKind.Object:
                var aProps = a.EnumerateObject().ToDictionary(p => p.Name, p => p, StringComparer.Ordinal);
                var bProps = b.EnumerateObject().ToDictionary(p => p.Name, p => p, StringComparer.Ordinal);
                foreach (var name in aProps.Keys.Except(bProps.Keys)) removed.Add(path + "." + name);
                foreach (var name in bProps.Keys.Except(aProps.Keys)) added.Add(path + "." + name);
                foreach (var name in aProps.Keys.Intersect(bProps.Keys))
                {
                    CompareElements(aProps[name].Value, bProps[name].Value, path + "." + name, added, removed, changed);
                }
                break;
            case JsonValueKind.Array:
                int len = Math.Max(a.GetArrayLength(), b.GetArrayLength());
                for (int i = 0; i < len; i++)
                {
                    if (i >= a.GetArrayLength()) { added.Add(path + "[" + i + "]"); continue; }
                    if (i >= b.GetArrayLength()) { removed.Add(path + "[" + i + "]"); continue; }
                    CompareElements(a[i], b[i], path + "[" + i + "]", added, removed, changed);
                }
                break;
            default:
                var aText = a.GetRawText();
                var bText = b.GetRawText();
                if (aText != bText)
                {
                    changed.Add((path, aText, bText));
                }
                break;
        }
    }
}
