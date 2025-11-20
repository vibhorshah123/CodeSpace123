using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using System.Net.Http;
using System.Text.Json;
using System.Text;
using System.Collections.Generic;
using System.Threading.Tasks;
using System;

namespace D365ComparissionTool;

public class Function1
{
    private readonly ILogger<Function1> _logger;

    public Function1(ILogger<Function1> logger)
    {
        _logger = logger;
    }

    [Function("Function1")]
    public async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.Function, "get", "post")] HttpRequest req)
    {
        string envUrl = req.Query["environmentUrl"].ToString();
        string bearer = req.Query["bearerToken"].ToString();
        string searchText = req.Query["text"].ToString();
        // Optional flags
        bool searchAllCategories = string.Equals(req.Query["searchAllCategories"], "true", StringComparison.OrdinalIgnoreCase);

        // If POST with JSON body, allow body override
        if (string.Equals(req.Method, "POST", StringComparison.OrdinalIgnoreCase))
        {
            try
            {
                using var reader = new System.IO.StreamReader(req.Body, Encoding.UTF8);
                var raw = await reader.ReadToEndAsync();
                if (!string.IsNullOrWhiteSpace(raw))
                {
                    var doc = JsonDocument.Parse(raw);
                    if (doc.RootElement.TryGetProperty("EnvironmentUrl", out var eu) && string.IsNullOrWhiteSpace(envUrl)) envUrl = eu.GetString() ?? string.Empty;
                    if (doc.RootElement.TryGetProperty("BearerToken", out var bt) && string.IsNullOrWhiteSpace(bearer)) bearer = bt.GetString() ?? string.Empty;
                    if (doc.RootElement.TryGetProperty("Text", out var st) && string.IsNullOrWhiteSpace(searchText)) searchText = st.GetString() ?? string.Empty;
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Failed to parse POST body");
            }
        }

        if (string.IsNullOrWhiteSpace(envUrl) || string.IsNullOrWhiteSpace(bearer) || string.IsNullOrWhiteSpace(searchText))
        {
            return new BadRequestObjectResult("Required: environmentUrl, bearerToken, text");
        }

        var correlationId = Guid.NewGuid();
        _logger.LogInformation("[FlowSearch] Start CorrelationId={CorrelationId} Env={Env} Text='{Text}'", correlationId, envUrl, searchText);

        try
        {
            var names = await FetchAndSearchFlowsAsync(envUrl, bearer, searchText, searchAllCategories, correlationId);
            var payload = new { MatchCount = names.Count, Names = names };
            var json = JsonSerializer.Serialize(payload, new JsonSerializerOptions { WriteIndented = true });
            return new OkObjectResult(json) { ContentTypes = { "application/json" } };
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "[FlowSearch] Failure CorrelationId={CorrelationId}", correlationId);
            return new ObjectResult("Error: " + ex.Message) { StatusCode = 500 };
        }
    }

    private async Task<List<string>> FetchAndSearchFlowsAsync(string envUrl, string bearerToken, string searchText, bool allCategories, Guid correlationId)
    {
        var names = new List<string>();
        using var client = new HttpClient();
        client.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", bearerToken);
        client.DefaultRequestHeaders.Accept.ParseAdd("application/json");
        string baseUrl = envUrl.TrimEnd('/') + "/api/data/v9.2/workflows";
        // Initial query (cloud flows category=5 unless user requested all)
        string filter = allCategories ? string.Empty : "&$filter=category eq 5";
        string url = baseUrl + "?$select=workflowid,name,uniquename,clientdata,description,category" + filter;
        int page = 0;
        var searchLower = searchText.ToLowerInvariant();
        while (!string.IsNullOrEmpty(url))
        {
            page++;
            _logger.LogDebug("[FlowSearch] Page={Page} URLLength={Len}", page, url.Length);
            using var req = new HttpRequestMessage(HttpMethod.Get, url);
            using var resp = await client.SendAsync(req);
            var content = await resp.Content.ReadAsStringAsync();
            if (!resp.IsSuccessStatusCode)
            {
                throw new HttpRequestException($"Workflow query failed status={resp.StatusCode} snippet={(content.Length > 300 ? content[..300] : content)}");
            }
            using var doc = JsonDocument.Parse(content);
            if (doc.RootElement.TryGetProperty("value", out var arr))
            {
                foreach (var wf in arr.EnumerateArray())
                {
                    string id = wf.TryGetProperty("workflowid", out var wid) ? wid.GetString() ?? string.Empty : string.Empty;
                    string name = wf.TryGetProperty("name", out var n) ? n.GetString() ?? string.Empty : string.Empty;
                    string unique = wf.TryGetProperty("uniquename", out var un) ? un.GetString() ?? string.Empty : string.Empty;
                    string clientData = wf.TryGetProperty("clientdata", out var cd) ? cd.GetString() ?? string.Empty : string.Empty;
                    string description = wf.TryGetProperty("description", out var desc) ? desc.GetString() ?? string.Empty : string.Empty;
                    bool hit = FieldHit(name, searchLower) || FieldHit(unique, searchLower) || FieldHit(description, searchLower) || ClientDataHit(clientData, searchLower);
                    if (hit && !string.IsNullOrWhiteSpace(name)) names.Add(name);
                }
            }
            url = doc.RootElement.TryGetProperty("@odata.nextLink", out var next) ? next.GetString() : null;
        }
        // Fallback: if no matches and we only searched category 5, broaden automatically
        if (names.Count == 0 && !allCategories)
        {
            _logger.LogInformation("[FlowSearch] No matches in category 5; broadening search to all workflow categories");
            return await FetchAndSearchFlowsAsync(envUrl, bearerToken, searchText, true, correlationId);
        }
        _logger.LogInformation("[FlowSearch] Completed Matches={Count}", names.Count);
        return names;

        static bool FieldHit(string? value, string searchLower) => !string.IsNullOrEmpty(value) && value.ToLowerInvariant().Contains(searchLower);

        static bool ClientDataHit(string clientData, string searchLower)
        {
            if (string.IsNullOrEmpty(clientData)) return false;
            // Direct string contains first
            if (clientData.ToLowerInvariant().Contains(searchLower)) return true;
            // Try base64 decode (some environments store encoded)
            try
            {
                if (IsBase64(clientData))
                {
                    var decoded = Encoding.UTF8.GetString(Convert.FromBase64String(clientData));
                    if (decoded.ToLowerInvariant().Contains(searchLower)) return true;
                    // If decoded JSON, search recursively
                    if (decoded.TrimStart().StartsWith("{"))
                    {
                        using var doc = JsonDocument.Parse(decoded);
                        if (JsonElementContains(doc.RootElement, searchLower)) return true;
                    }
                }
                else if (clientData.TrimStart().StartsWith("{"))
                {
                    using var doc = JsonDocument.Parse(clientData);
                    if (JsonElementContains(doc.RootElement, searchLower)) return true;
                }
            }
            catch { }
            return false;
        }

        static bool IsBase64(string s)
        {
            s = s.Trim();
            if (s.Length < 16 || s.Length % 4 != 0) return false;
            Span<byte> buffer = new Span<byte>(new byte[s.Length]);
            return Convert.TryFromBase64String(s, buffer, out _);
        }

        static bool JsonElementContains(JsonElement el, string searchLower)
        {
            switch (el.ValueKind)
            {
                case JsonValueKind.String:
                    return el.GetString()?.ToLowerInvariant().Contains(searchLower) == true;
                case JsonValueKind.Number:
                    return el.GetRawText().ToLowerInvariant().Contains(searchLower);
                case JsonValueKind.True:
                case JsonValueKind.False:
                    return el.GetRawText().ToLowerInvariant().Contains(searchLower);
                case JsonValueKind.Object:
                    foreach (var p in el.EnumerateObject())
                    {
                        if (p.Name.ToLowerInvariant().Contains(searchLower)) return true;
                        if (JsonElementContains(p.Value, searchLower)) return true;
                    }
                    return false;
                case JsonValueKind.Array:
                    foreach (var item in el.EnumerateArray()) if (JsonElementContains(item, searchLower)) return true;
                    return false;
                default:
                    return el.GetRawText().ToLowerInvariant().Contains(searchLower);
            }
        }
    }
}