/// <summary>
/// The Etherpad-specific code used by the WOPI bridge.
/// </summary>
class Etherpad(BridgeHttp http, WopiClient wopi, ILogger log) :
    BridgeApp(http, wopi, log)
{
    string groupId = "";

    /// <summary>
    /// Creates a general group to attach all pads.
    /// </summary>
    protected override async Task Probe()
    {
        var result = await ApiCall(
            "createGroupIfNotExistsFor",
            [("groupMapper", "1")]);
        groupId = (string?) result["data"]?["groupID"] ?? throw new AppFailure("Error response from Etherpad");
        Log.LogInformation("Got Etherpad global groupid {groupid}", groupId);
    }

    /// <summary>
    /// Generic method to call the Etherpad REST API.
    /// </summary>
    async Task<JsonNode> ApiCall(string method, List<(string Name, string Value)> parameters, string? acctok = null, bool raiseOnNonZeroCode = true)
    {
        parameters.Add(("apikey", ApiKey ?? ""));
        var tokenForLog = acctok?.Short() ?? "N/A";
        JsonNode result;
        try
        {
            using var response = await Http.Follow.PostAsync($"{AppUrl}/api/1/{method}?{Query(parameters.ToArray())}", null);
            var content = await response.Content.ReadAsByteArrayAsync();
            if (response.StatusCode != HttpStatusCode.OK)
            {
                Log.LogError("Failed to call Etherpad {method} {token} {response}", method, tokenForLog, $"{(int) response.StatusCode}: {Body(content)}");
                throw new AppFailure("Failed to connect to Etherpad");
            }

            result = JsonNode.Parse(content) ?? throw new AppFailure("Error response from Etherpad");
        }
        catch (Exception exception) when (IsConnectionError(exception) || exception is JsonException)
        {
            Log.LogError("Exception raised attempting to connect to Etherpad {method} {exception}", method, exception.Message);
            throw new AppFailure("Failed to connect to Etherpad", exception);
        }

        if ((int?) result["code"] != 0 &&
            raiseOnNonZeroCode)
        {
            Log.LogError("Error response from Etherpad {method} {token} {response}", method, tokenForLog, (string?) result["message"]);
            throw new AppFailure("Error response from Etherpad");
        }

        Log.LogDebug("Called Etherpad API {method} {token} {result}", method, tokenForLog, result.ToJsonString());
        return result;
    }

    public override async Task<string> GetRedirectUrl(ViewMode viewMode, string wopiSrc, string acctok, string docId, string filename, string displayName, string? revaToken)
    {
        if (viewMode is ViewMode.ReadOnly or ViewMode.ViewOnly)
        {
            // for read-only mode generate a read-only link
            var result = await ApiCall("getReadOnlyID", [("padID", docId)], acctok);
            return $"{AppExtUrl}/p/{(string?) result["data"]?["readOnlyID"]}?userName={UrlCodec.QuotePlus(displayName)}";
        }

        // pass to Etherpad the required metadata for the save webhook
        try
        {
            var query = Query(("padID", docId), ("wopiSrc", UrlCodec.QuotePlus(wopiSrc)), ("accessToken", acctok), ("apikey", ApiKey ?? ""));
            using var response = await Http.Follow.PostAsync($"{AppUrl}/setEFSSMetadata?{query}", null);
            var content = await response.Content.ReadAsByteArrayAsync();
            if (response.StatusCode != HttpStatusCode.OK ||
                (int?) ParseOrNull(content)?["code"] != 0)
            {
                Log.LogError("Failed to call Etherpad {method} {token} {response}", "setEFSSMetadata", acctok.Short(), $"{(int) response.StatusCode}: {Body(content).Replace('"', '\'')}");
                throw new AppFailure("Error response from Etherpad");
            }

            Log.LogDebug("Called Etherpad {method} {token}", "setEFSSMetadata", acctok.Short());
        }
        catch (Exception exception) when (IsConnectionError(exception))
        {
            Log.LogError("Exception raised attempting to connect to Etherpad {method} {exception}", "setEFSSMetadata", exception.Message);
            throw new AppFailure("Failed to connect to Etherpad", exception);
        }

        // return the URL to the pad for editing (a PREVIEW viewmode is not supported)
        return $"{AppExtUrl}/p/{docId}?userName={UrlCodec.QuotePlus(displayName)}";
    }

    static JsonNode? ParseOrNull(byte[] content)
    {
        try
        {
            return JsonNode.Parse(content);
        }
        catch (JsonException)
        {
            return null;
        }
    }

    // Cloud storage to Etherpad

    /// <summary>
    /// Copies the document from storage to Etherpad.
    /// </summary>
    public override async Task<BridgeLock> LoadFromStorage(JsonNode fileMetadata, string wopiSrc, string acctok, string? docId)
    {
        // WOPI GetFile
        var fileResponse = await Wopi.Request(wopiSrc, acctok, HttpMethod.Get, contents: []);
        if (fileResponse.Status != StatusCodes.Status200OK)
        {
            throw new AppFailure($"Unable to fetch file from storage, got HTTP {fileResponse.Status}");
        }

        var epFile = fileResponse.Content;
        if (epFile.Length == 0)
        {
            epFile = "{}"u8.ToArray();
        }

        if (docId == null)
        {
            docId = WopiUtils.RandomString(20);
            Log.LogDebug("Generated random padID for read-only document {padid} {token}", docId, acctok.Short());
        }

        // first drop previous pad if it exists
        await ApiCall("deletePad", [("padID", docId)], acctok, raiseOnNonZeroCode: false);
        // create pad with the given docid as name
        await ApiCall("createGroupPad", [("groupID", groupId), ("padName", docId), ("text", "placeholder")], acctok, raiseOnNonZeroCode: false);
        try
        {
            // push content: a .etherpad file is imported as raw (JSON) content
            using var form = new MultipartFormDataContent();
            var file = new ByteArrayContent(epFile);
            file.Headers.ContentType = new("application/json");
            form.Add(file, "file", $"{docId}.etherpad");
            using var response = await Http.Follow.PostAsync($"{AppUrl}/p/{docId}/import?{Query(("apikey", ApiKey ?? ""))}", form);
            if (response.StatusCode != HttpStatusCode.OK)
            {
                var content = await response.Content.ReadAsByteArrayAsync();
                Log.LogError("Unable to push document to Etherpad {token} {padid} {response}", acctok.Short(), docId, $"{(int) response.StatusCode}: {Body(content)}");
                throw new AppFailure("Error response from Etherpad");
            }

            Log.LogInformation("Pushed document to Etherpad {padid} {token}", docId, acctok.Short());
        }
        catch (Exception exception) when (IsConnectionError(exception))
        {
            Log.LogError("Exception raised attempting to connect to Etherpad {method} {exception}", "import", exception.Message);
            throw new AppFailure("Failed to connect to Etherpad", exception);
        }

        // generate and return a WOPI lock structure for this document
        return WopiClient.GenerateLock(docId, fileMetadata, epFile, acctok, false);
    }

    // Etherpad to cloud storage

    /// <summary>
    /// Fetches a given document from Etherpad, throws <see cref="AppFailure"/> in case of errors.
    /// </summary>
    public override async Task<byte[]> Fetch(BridgeLock wopiLock, string acctok)
    {
        try
        {
            // this operation does not use the API (and it is NOT protected by the API key!), so we use a plain GET
            using var response = await Http.Follow.GetAsync($"{AppUrl}/p{wopiLock.Doc}/export/etherpad");
            var content = await response.Content.ReadAsByteArrayAsync();
            if (response.StatusCode != HttpStatusCode.OK)
            {
                Log.LogError("Unable to fetch document from Etherpad {token} {response}", acctok.Short(), $"{(int) response.StatusCode}: {Body(content, 50)}");
                throw new AppFailure();
            }

            return content;
        }
        catch (Exception exception) when (IsConnectionError(exception))
        {
            Log.LogError("Exception raised attempting to connect to Etherpad {exception}", exception.Message);
            throw new AppFailure("", exception);
        }
    }

    /// <summary>
    /// Copies the document from Etherpad back to storage.
    /// </summary>
    public override async Task<BridgeReply> SaveToStorage(string wopiSrc, string acctok, bool isClose, BridgeLock wopiLock)
    {
        // get document from Etherpad
        byte[] epFile;
        try
        {
            Log.LogInformation("Fetching file from Etherpad {isclose} {appurl} {token}", isClose, $"{AppUrl}/p{wopiLock.Doc}", acctok.Short());
            epFile = await Fetch(wopiLock, acctok);
        }
        catch (AppFailure)
        {
            // return a non-fatal error
            return WopiClient.Reply("File not saved, error in fetching document from Etherpad. Will try again later", StatusCodes.Status424FailedDependency);
        }

        if (isClose &&
            Wopi.CheckForNoChanges(epFile, wopiLock, acctok))
        {
            return new("{}", StatusCodes.Status202Accepted);
        }

        // WOPI PutFile
        var response = await Wopi.Request(
            wopiSrc,
            acctok,
            HttpMethod.Post,
            epFile,
            new()
            {
                ["X-WOPI-Lock"] = wopiLock.ToJson()
            });
        var reply = Wopi.HandlePutFile("PutFile", wopiSrc, response);
        if (reply != null)
        {
            return reply;
        }

        return await Wopi.RefreshDigestAndLock(wopiSrc, acctok, wopiLock, epFile);
    }
}
