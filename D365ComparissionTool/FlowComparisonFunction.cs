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
    string Notes);

public class FlowComparisonFunction
{
    private readonly ILogger _logger;
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true
    };

    public FlowComparisonFunction(ILogger<FlowComparisonFunction> logger) => _logger = logger;

    [Function("CompareFlows")] // HTTP trigger for flow comparison
    public async Task<HttpResponseData> Run([HttpTrigger(AuthorizationLevel.Function, "post", Route = "CompareFlows")] HttpRequestData req, FunctionContext ctx)
    {
        var correlationId = Guid.NewGuid();
        _logger.LogInformation("[CompareFlows] Start CorrelationId={CorrelationId}", correlationId);
        FlowComparisonRequestPayload payload;
        try
        {
            string body = await new StreamReader(req.Body).ReadToEndAsync();
            payload = JsonSerializer.Deserialize<FlowComparisonRequestPayload>(body, JsonOptions) ?? new FlowComparisonRequestPayload();
            _logger.LogDebug("[CompareFlows:{CorrelationId}] Raw body length={BodyLength}", correlationId, body.Length);
        }
        catch (Exception ex)
        {
            var bad = req.CreateResponse(System.Net.HttpStatusCode.BadRequest);
            await bad.WriteStringAsync("Invalid JSON payload: " + ex.Message);
            return bad;
        }

        string mode = payload.Mode?.Trim().ToLowerInvariant() ?? "all";
        if (string.IsNullOrWhiteSpace(payload.Env))
        {
            var bad = req.CreateResponse(System.Net.HttpStatusCode.BadRequest);
            await bad.WriteStringAsync("Missing 'env' value.");
            return bad;
        }

        string env1 = NormalizeEnvHost(payload.Env!);
        string? env2 = string.IsNullOrWhiteSpace(payload.Env2) ? null : NormalizeEnvHost(payload.Env2!);
        _logger.LogInformation("[CompareFlows:{CorrelationId}] Mode={Mode} Env1={Env1} Env2={Env2} MatchBy={MatchBy} IncludeDiffDetails={IncludeDiffDetails}", correlationId, mode, env1, env2 ?? "(none)", payload.MatchBy ?? "name", payload.IncludeDiffDetails);

        // Primary token now expected in body as token1; fallback to Authorization header if omitted.
        string token1 = (payload.Token1 ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(token1)) token1 = ExtractBearerToken(req).Trim();
        if (string.IsNullOrWhiteSpace(token1))
        {
            var bad = req.CreateResponse(System.Net.HttpStatusCode.BadRequest);
            await bad.WriteStringAsync("token1 (primary environment bearer) is required in body or Authorization header.");
            return bad;
        }
        string? token2 = payload.Token2?.Trim();
        if (env2 != null && string.IsNullOrWhiteSpace(token2))
        {
            var bad = req.CreateResponse(System.Net.HttpStatusCode.BadRequest);
            await bad.WriteStringAsync("token2 required when env2 provided.");
            return bad;
        }

        // Validate modes
        if (mode is "single")
        {
            if (string.IsNullOrWhiteSpace(payload.FlowId) && string.IsNullOrWhiteSpace(payload.FlowName))
            {
                var bad = req.CreateResponse(System.Net.HttpStatusCode.BadRequest);
                await bad.WriteStringAsync("single mode requires flowId or flowName");
                return bad;
            }
        }
        else if (mode is "list")
        {
            if (payload.Flows == null || payload.Flows.Count == 0)
            {
                var bad = req.CreateResponse(System.Net.HttpStatusCode.BadRequest);
                await bad.WriteStringAsync("list mode requires 'flows' array of names or GUIDs");
                return bad;
            }
        }

        var ignoreKeysConfig = Environment.GetEnvironmentVariable("NORMALIZE_IGNORE_KEYS") ?? string.Empty;
        var ignoreKeys = ignoreKeysConfig.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Select(k => k.Trim()).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        if (ignoreKeys.Count == 0)
        {
            ignoreKeys.AddRange(new[] { "connectionReferences", "runtimeConfiguration", "lastModified", "createdTime", "modifiedTime", "etag", "trackedProperties" });
        }

        var env1Flows = new List<FlowSnapshot>();
        var env2Flows = new List<FlowSnapshot>();
        string notes = string.Empty;
        var fetchStart = DateTime.UtcNow;

        try
        {
            if (mode == "all")
            {
                env1Flows = await FetchAndPrepareAllFlows(env1, token1, ignoreKeys, correlationId);
                if (env2 != null) env2Flows = await FetchAndPrepareAllFlows(env2, token2!, ignoreKeys, correlationId);
            }
            else if (mode == "single")
            {
                string identifier = payload.FlowId ?? payload.FlowName!;
                _logger.LogInformation("[CompareFlows:{CorrelationId}] Single identifier={Identifier}", correlationId, identifier);
                var snap1 = await FetchSingleFlow(env1, token1, identifier, ignoreKeys, correlationId);
                if (snap1 != null) env1Flows.Add(snap1);
                if (env2 != null)
                {
                    var snap2 = await FetchSingleFlow(env2, token2!, identifier, ignoreKeys, correlationId, payload.MatchBy);
                    if (snap2 != null) env2Flows.Add(snap2);
                }
            }
            else if (mode == "list")
            {
                _logger.LogInformation("[CompareFlows:{CorrelationId}] List mode count={Count}", correlationId, payload.Flows!.Count);
                foreach (var f in payload.Flows!)
                {
                    var s = await FetchSingleFlow(env1, token1, f, ignoreKeys, correlationId);
                    if (s != null) env1Flows.Add(s);
                }
                if (env2 != null)
                {
                    foreach (var f in payload.Flows!)
                    {
                        var s = await FetchSingleFlow(env2, token2!, f, ignoreKeys, correlationId);
                        if (s != null) env2Flows.Add(s);
                    }
                }
            }
            else
            {
                notes = "Unknown mode; defaulted to all.";
                env1Flows = await FetchAndPrepareAllFlows(env1, token1, ignoreKeys, correlationId);
                if (env2 != null) env2Flows = await FetchAndPrepareAllFlows(env2, token2!, ignoreKeys, correlationId);
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "[CompareFlows:{CorrelationId}] Fetch failure", correlationId);
            var errResp = req.CreateResponse(System.Net.HttpStatusCode.InternalServerError);
            await errResp.WriteStringAsync("Flow fetch failed: " + ex.Message);
            return errResp;
        }
        var fetchDurationMs = (int)(DateTime.UtcNow - fetchStart).TotalMilliseconds;
        _logger.LogInformation("[CompareFlows:{CorrelationId}] Fetch+Prepare completed Env1Count={Env1Count} Env2Count={Env2Count} DurationMs={Duration}", correlationId, env1Flows.Count, env2Flows.Count, fetchDurationMs);

        var comparisons = new List<FlowComparisonResult>();
        if (env2 != null)
        {
            string matchBy = (payload.MatchBy?.Trim().ToLowerInvariant()) switch
            {
                "id" => "id",
                _ => "name"
            };
            var env2LookupByName = env2Flows.GroupBy(f => f.Name, StringComparer.OrdinalIgnoreCase).ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);
            var env2LookupById = env2Flows.ToDictionary(f => f.FlowId, f => f);

            foreach (var f1 in env1Flows)
            {
                FlowSnapshot? f2 = null;
                if (matchBy == "id") env2LookupById.TryGetValue(f1.FlowId, out f2);
                else env2LookupByName.TryGetValue(f1.Name, out f2);
                if (f2 == null)
                {
                    comparisons.Add(new FlowComparisonResult(f1.Name, f1, null, false, null, null, "missing_in_env2"));
                    _logger.LogDebug("[CompareFlows:{CorrelationId}] MissingInEnv2 Name={Name} Id={Id}", correlationId, f1.Name, f1.FlowId);
                    continue;
                }
                if (!string.IsNullOrEmpty(f1.Error) || !string.IsNullOrEmpty(f2.Error))
                {
                    comparisons.Add(new FlowComparisonResult(f1.Name, f1, f2, false, null, null, "error"));
                    _logger.LogWarning("[CompareFlows:{CorrelationId}] Error snapshot Name={Name} Env1Error={E1} Env2Error={E2}", correlationId, f1.Name, f1.Error, f2.Error);
                    continue;
                }
                bool identical = string.Equals(f1.Hash, f2.Hash, StringComparison.OrdinalIgnoreCase);
                FlowComparisonDiff? diff = null;
                List<FlowActionDifference>? actionDiffs = null;
                if (!identical && payload.IncludeDiffDetails)
                {
                    diff = ComputeDiff(f1.CanonicalJson, f2.CanonicalJson);
                    _logger.LogDebug("[CompareFlows:{CorrelationId}] Diff Name={Name} Added={Added} Removed={Removed} Changed={Changed}", correlationId, f1.Name, diff.Added.Count, diff.Removed.Count, diff.Changed.Count);
                    actionDiffs = ComputeActionDifferences(f1.CanonicalJson, f2.CanonicalJson);
                    if (actionDiffs != null && actionDiffs.Count > 0)
                    {
                        _logger.LogDebug("[CompareFlows:{CorrelationId}] ActionDiffs Name={Name} NonMatchingActions={Count}", correlationId, f1.Name, actionDiffs.Count);
                    }
                }
                comparisons.Add(new FlowComparisonResult(f1.Name, f1, f2, identical, diff, actionDiffs, identical ? "identical" : "different"));
            }
        }

        int identicalCount = comparisons.Count(c => c.Identical); // retained for logging only
        int differentCount = comparisons.Count(c => c.Status == "different"); // retained for logging only
        int missingCount = comparisons.Count(c => c.Status == "missing_in_env2");
        int errorCount = comparisons.Count(c => c.Status == "error");
        var identicalFlows = comparisons.Where(c => c.Identical).Select(c => c.Name).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        var nonIdenticalFlows = comparisons.Where(c => c.Status == "different").Select(c => c.Name).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        var responseObj = new FlowComparisonResponse(env1Flows, env2Flows, comparisons, identicalFlows, nonIdenticalFlows, missingCount, errorCount, notes);
        var resp = req.CreateResponse(System.Net.HttpStatusCode.OK);
        resp.Headers.Add("Content-Type", "application/json");
        await resp.WriteStringAsync(JsonSerializer.Serialize(responseObj, new JsonSerializerOptions { WriteIndented = true }));
        _logger.LogInformation("[CompareFlows] Completed CorrelationId={CorrelationId} Env1Flows={Env1Count} Env2Flows={Env2Count} Identical={Identical} Different={Different} MissingEnv2={Missing} Errors={Errors}", correlationId, env1Flows.Count, env2Flows.Count, identicalCount, differentCount, missingCount, errorCount);
        return resp;
    }

    private static string ExtractBearerToken(HttpRequestData req)
    {
        if (req.Headers.TryGetValues("Authorization", out var authValues))
        {
            var bearer = authValues.FirstOrDefault(v => v.StartsWith("Bearer ", StringComparison.OrdinalIgnoreCase));
            if (bearer != null) return bearer.Substring(7).Trim();
        }
        return string.Empty;
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
        // Add headers required by Dataverse for some endpoints to avoid 400/415 issues
        if (!client.DefaultRequestHeaders.Contains("OData-MaxVersion"))
            client.DefaultRequestHeaders.Add("OData-MaxVersion", "4.0");
        if (!client.DefaultRequestHeaders.Contains("OData-Version"))
            client.DefaultRequestHeaders.Add("OData-Version", "4.0");
        if (!client.DefaultRequestHeaders.Contains("User-Agent"))
            client.DefaultRequestHeaders.Add("User-Agent", "D365ComparisonTool/1.0");
        client.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);
        client.DefaultRequestHeaders.Accept.ParseAdd("application/json");
        // Use Dataverse workflows entity (holds cloud flow definitions) selecting workflowid,name,clientdata
        // Category 5 corresponds to modern (cloud) flows; include filter when not doing single workflowid lookup by GUID
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
