/// <summary>
/// The lock held by the bridge on a document, in the format { doc, fn, dig, tocl }, where tocl maps the short
/// tokens of all known sessions to whether they were closed, and dig is a SHA1 digest of the content for later
/// checks when the file gets modified.
/// </summary>
class BridgeLock
{
    [JsonPropertyName("doc")]
    public required string Doc { get; set; }

    [JsonPropertyName("fn")]
    public required string FileName { get; set; }

    [JsonPropertyName("dig")]
    public required string Digest { get; set; }

    [JsonPropertyName("tocl")]
    public required Dictionary<string, bool> ToClose { get; set; }

    public string ToJson() =>
        Json.Serialize(this);

    public BridgeLock Clone() =>
        Json.Deserialize<BridgeLock>(ToJson())!;
}

/// <summary>
/// A response of the bridge to a save request: a JSON-formatted message and an HTTP status code.
/// </summary>
record BridgeReply(string Message, int Status);

/// <summary>
/// A WOPI response, fully read.
/// </summary>
class WopiResponse(int status, Dictionary<string, string> headers, byte[] content)
{
    public int Status { get; } = status;

    public byte[] Content { get; } = content;

    public string? Header(string name) =>
        headers.GetValueOrDefault(name);

    public JsonNode? Json() =>
        JsonNode.Parse(Content);
}

/// <summary>
/// An invalid or missing WOPI lock.
/// </summary>
class InvalidLockException(string message) :
    Exception(message);

/// <summary>
/// A set of WOPI client functions for the WOPI bridge service.
/// </summary>
class WopiClient(HttpClient http, bool sslVerify, ILogger log)
{
    /// <summary>
    /// Consistently json-ifies a given message and passes a delay to dismiss it: a 0 delay means the user has to
    /// click on it to dismiss it, good for longer messages. This is part of the EFSS webhook specs implemented by
    /// the bridged apps.
    /// </summary>
    public static string Jsonify(string message)
    {
        var delay = message.Length > 60 ? 0 : 0.5 + message.Length / 20.0;
        return $$"""{"message": "{{JsonEncodedText.Encode(message, System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping)}}", "delay": "{{delay.ToString("F1", CultureInfo.InvariantCulture)}}"}""";
    }

    public static BridgeReply Reply(string message, int status) =>
        new(Jsonify(message), status);

    /// <summary>
    /// Executes a WOPI request with the given parameters and headers. Failures to contact the server are
    /// reported as a 500 response.
    /// </summary>
    public async Task<WopiResponse> Request(string wopiSrc, string acctok, HttpMethod method, byte[]? contents = null, Dictionary<string, string>? headers = null)
    {
        var isPutRelative = headers != null &&
                            headers.Any(_ => string.Equals(_.Key, "X-WOPI-Override", StringComparison.OrdinalIgnoreCase) && _.Value == "PUT_RELATIVE");
        var wopiUrl = contents != null && !isPutRelative ? $"{wopiSrc}/contents" : wopiSrc;
        log.LogDebug("Calling WOPI {url} {headers} {acctok} {ssl}", wopiUrl, headers == null ? "None" : Json.Serialize(headers), acctok.Short(), sslVerify);
        try
        {
            using var request = new HttpRequestMessage(method, $"{wopiUrl}?access_token={acctok}");
            if (method == HttpMethod.Post)
            {
                request.Content = new ByteArrayContent(contents ?? []);
            }

            foreach (var (name, value) in headers ?? [])
            {
                request.Headers.TryAddWithoutValidation(name, value);
            }

            using var response = await http.SendAsync(request);
            var responseHeaders = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var (name, values) in response.Headers.Concat(response.Content.Headers))
            {
                responseHeaders[name] = string.Join(", ", values);
            }

            return new((int) response.StatusCode, responseHeaders, await response.Content.ReadAsByteArrayAsync());
        }
        catch (Exception exception) when (exception is HttpRequestException or TaskCanceledException or IOException)
        {
            log.LogError("Unable to contact WOPI {wopiurl} {acctok} {response}", wopiUrl, acctok, exception.Message);
            return new(StatusCodes.Status500InternalServerError, [], []);
        }
    }

    public static string Digest(byte[] content) =>
        Convert.ToHexStringLower(SHA1.HashData(content));

    /// <summary>
    /// Returns a lock to be used as WOPI lock.
    /// </summary>
    public static BridgeLock GenerateLock(string docId, JsonNode fileMetadata, byte[]? content, string acctok, bool isClose) =>
        new()
        {
            Doc = $"/{docId.Trim('/')}",
            FileName = (string) fileMetadata["BaseFileName"]!,
            Digest = content is { Length: > 0 } ? Digest(content) : "dirty",
            ToClose = new()
            {
                [acctok.Short()] = isClose
            }
        };

    /// <summary>
    /// Returns true if the content did not change from the digest stored in the lock.
    /// </summary>
    public bool CheckForNoChanges(byte[] content, BridgeLock wopiLock, string acctokForLog)
    {
        // so far the file was not touched: let's validate the content
        if (wopiLock.Digest != "dirty" &&
            Digest(content) == wopiLock.Digest)
        {
            log.LogInformation("File unchanged, skipping save {token}", acctokForLog.Short());
            return true;
        }

        return false;
    }

    /// <summary>
    /// Returns the currently held WOPI lock, or throws <see cref="InvalidLockException"/> otherwise.
    /// The message of the exception is the HTTP status when the request failed.
    /// </summary>
    public async Task<BridgeLock> GetLock(string wopiSrc, string acctok)
    {
        var response = await Request(
            wopiSrc,
            acctok,
            HttpMethod.Post,
            headers: new()
            {
                ["X-Wopi-Override"] = "GET_LOCK"
            });
        if (response.Status != StatusCodes.Status200OK)
        {
            // lock got lost or any other error
            throw new InvalidLockException(response.Status.ToString(CultureInfo.InvariantCulture));
        }

        try
        {
            // the lock is expected to be a JSON dict, see GenerateLock()
            return Json.Deserialize<BridgeLock>(response.Header("X-WOPI-Lock") ?? throw new KeyNotFoundException("'X-WOPI-Lock'")) ??
                   throw new JsonException("null lock");
        }
        catch (Exception exception) when (exception is JsonException or KeyNotFoundException)
        {
            log.LogWarning("Missing or malformed WOPI lock {exception}", $"{exception.GetType().Name}: {exception.Message}");
            throw new InvalidLockException(exception.Message);
        }
    }

    /// <summary>
    /// Helper function for relock to generate the old and new lock structures.
    /// </summary>
    static (Dictionary<string, string> Headers, BridgeLock NewLock) GetHeadersForRelock(string acctok, BridgeLock wopiLock, string? digest, Dictionary<string, bool>? toClose)
    {
        var newLock = wopiLock.Clone();
        if (toClose != null)
        {
            // we got the full 'toclose' dict, push it as is
            newLock.ToClose = new(toClose);
        }
        else if (!wopiLock.ToClose.ContainsKey(acctok.Short()))
        {
            // if missing, just append the short token to the 'toclose' dict, similarly to the openfiles map
            newLock.ToClose[acctok.Short()] = false;
        }

        if (digest != null &&
            wopiLock.Digest != digest)
        {
            newLock.Digest = digest;
        }

        var headers = new Dictionary<string, string>
        {
            ["X-Wopi-Override"] = "LOCK",
            ["X-WOPI-OldLock"] = wopiLock.ToJson(),
            ["X-WOPI-Lock"] = newLock.ToJson()
        };
        return (headers, newLock);
    }

    /// <summary>
    /// Refreshes an existing WOPI lock. Returns the new lock if successful, throws <see cref="InvalidLockException"/> otherwise.
    /// </summary>
    public async Task<BridgeLock> RefreshLock(string wopiSrc, string acctok, BridgeLock wopiLock, string? digest = null, Dictionary<string, bool>? toClose = null)
    {
        var (headers, newLock) = GetHeadersForRelock(acctok, wopiLock, digest, toClose);
        var response = await Request(wopiSrc, acctok, HttpMethod.Post, headers: headers);
        if (response.Status == StatusCodes.Status200OK)
        {
            return newLock;
        }

        if (response.Status == StatusCodes.Status409Conflict)
        {
            // we have a race condition, another thread has updated the lock before us
            log.LogWarning("Got conflict in refreshing lock, retrying {url}", wopiSrc);
            BridgeLock currentLock;
            try
            {
                currentLock = Json.Deserialize<BridgeLock>(response.Header("X-WOPI-Lock") ?? "") ?? throw new JsonException("null lock");
            }
            catch (JsonException exception)
            {
                log.LogError("Got unresolvable conflict in RefreshLock {url} {previouslock} {error}", wopiSrc, response.Header("X-WOPI-Lock"), exception.Message);
                throw new InvalidLockException("Found existing malformed lock on refreshlock");
            }

            if (toClose != null)
            {
                // merge toclose token lists
                foreach (var (shortToken, isClosed) in currentLock.ToClose)
                {
                    toClose[shortToken] = isClosed || toClose.GetValueOrDefault(shortToken);
                }
            }

            if (digest != null)
            {
                wopiLock.Digest = currentLock.Digest;
            }

            // retry with the newly got lock
            (headers, newLock) = GetHeadersForRelock(acctok, wopiLock, digest, toClose);
            response = await Request(wopiSrc, acctok, HttpMethod.Post, headers: headers);
            if (response.Status == StatusCodes.Status200OK)
            {
                return newLock;
            }
        }

        log.LogError("Calling WOPI RefreshLock failed {url} {response} {reason}", wopiSrc, response.Status, response.Header("X-WOPI-LockFailureReason"));
        throw new InvalidLockException("Failed to refresh the lock");
    }

    /// <summary>
    /// Following a save operation, recomputes a digest for the given content and refreshes the lock.
    /// </summary>
    public async Task<BridgeReply> RefreshDigestAndLock(string wopiSrc, string acctok, BridgeLock wopiLock, byte[] content)
    {
        string? digest = null;
        if (wopiLock.Digest == "dirty")
        {
            digest = Digest(content);
        }

        try
        {
            wopiLock = await RefreshLock(wopiSrc, acctok, wopiLock, digest);
            log.LogInformation("Save completed {filename} {dig} {token}", wopiLock.FileName, digest, acctok.Short());
            return Reply("File saved successfully", StatusCodes.Status200OK);
        }
        catch (InvalidLockException)
        {
            return Reply("File saved, but failed to refresh lock", StatusCodes.Status500InternalServerError);
        }
    }

    /// <summary>
    /// Relocks a given document and returns a valid WOPI lock, or throws <see cref="InvalidLockException"/>.
    /// </summary>
    public async Task<BridgeLock> Relock(string wopiSrc, string acctok, string docId, bool isClose)
    {
        // first get again the file metadata
        var response = await Request(wopiSrc, acctok, HttpMethod.Get);
        if (response.Status != StatusCodes.Status200OK)
        {
            log.LogWarning("Session expired or file renamed when attempting to relock it {response} {docid} {token}", response.Status, docId, acctok.Short());
            throw new InvalidLockException("Session expired, please refresh this page");
        }

        var fileMetadata = response.Json()!;

        // lock the file again: we assume we are alone as the previous lock had been released
        var wopiLock = GenerateLock(docId, fileMetadata, null, acctok, isClose);
        var lockHeaders = new Dictionary<string, string>
        {
            ["X-WOPI-Lock"] = wopiLock.ToJson(),
            ["X-WOPI-Override"] = "REFRESH_LOCK",
            // this is an extension of the Lock API
            ["X-WOPI-Validate-Target"] = "True"
        };
        response = await Request(wopiSrc, acctok, HttpMethod.Post, headers: lockHeaders);
        if (response.Status == StatusCodes.Status409Conflict)
        {
            log.LogWarning("Got conflict in relocking the file {response} {token} {reason}", response.Status, acctok.Short(), response.Header("X-WOPI-LockFailureReason"));
            throw new InvalidLockException("The file was modified externally, please refresh this page to get its current version");
        }

        if (response.Status != StatusCodes.Status200OK)
        {
            log.LogWarning("Failed to relock the file {response} {token} {reason}", response.Status, acctok.Short(), response.Header("X-WOPI-LockFailureReason"));
            throw new InvalidLockException("Failed to relock the file on save, please refresh this page");
        }

        // relock was successful, return lock: along with docids univocally associated to files (WOPISrc's),
        // we are sure no other updates could have been missed
        return wopiLock;
    }

    /// <summary>
    /// Deals with conflicts or errors following a PutFile/PutRelative request.
    /// </summary>
    public BridgeReply? HandlePutFile(string wopiCall, string wopiSrc, WopiResponse response)
    {
        if (response.Status == StatusCodes.Status409Conflict)
        {
            // this is typically a user issue, return 500 and stop further editing
            log.LogWarning("Conflict when calling WOPI {wopicall} {url} {reason}", wopiCall, wopiSrc, response.Header("X-WOPI-LockFailureReason"));
            return Reply($"Error saving the file. {response.Header("X-WOPI-LockFailureReason")}", StatusCodes.Status500InternalServerError);
        }

        if (response.Status == StatusCodes.Status500InternalServerError)
        {
            // hopefully this is transient and the server has kept a local copy for later recovery
            log.LogError("Calling WOPI {wopicall} failed, will retry {url} {response}", wopiCall, wopiSrc, response.Status);
            return Reply("Error saving the file, will try again", StatusCodes.Status424FailedDependency);
        }

        if (response.Status != StatusCodes.Status200OK)
        {
            // any other error is considered also fatal
            log.LogError("Calling WOPI {wopicall} failed {url} {response}", wopiCall, wopiSrc, response.Status);
            return Reply("Error saving the file, please contact support", StatusCodes.Status500InternalServerError);
        }

        return null;
    }

    /// <summary>
    /// Saves a given document with an alternate name by using WOPI PutRelative.
    /// </summary>
    public async Task<BridgeReply> SaveAs(string wopiSrc, string acctok, BridgeLock wopiLock, string targetName, byte[] content)
    {
        var putRelativeHeaders = new Dictionary<string, string>
        {
            ["X-WOPI-Lock"] = wopiLock.ToJson(),
            ["X-WOPI-Override"] = "PUT_RELATIVE",
            // SuggestedTarget to not overwrite a possibly existing file
            ["X-WOPI-SuggestedTarget"] = targetName
        };
        var response = await Request(wopiSrc, acctok, HttpMethod.Post, content, putRelativeHeaders);
        var reply = HandlePutFile("PutRelative", wopiSrc, response);
        if (reply != null)
        {
            return reply;
        }

        // use the new file's metadata from PutRelative to remove the previous file: we can do that only on close
        // because we need to keep using the current wopisrc/acctok until the session is alive in the app
        var newName = (string?) response.Json()?["Name"];
        // unlock and delete original file
        response = await Request(
            wopiSrc,
            acctok,
            HttpMethod.Post,
            headers: new()
            {
                ["X-WOPI-Lock"] = wopiLock.ToJson(),
                ["X-Wopi-Override"] = "UNLOCK"
            });
        if (response.Status != StatusCodes.Status200OK)
        {
            log.LogWarning("Failed to unlock the previous file {token} {response}", acctok.Short(), response.Status);
        }
        else
        {
            response = await Request(
                wopiSrc,
                acctok,
                HttpMethod.Post,
                headers: new()
                {
                    ["X-Wopi-Override"] = "DELETE"
                });
            if (response.Status != StatusCodes.Status200OK)
            {
                log.LogWarning("Failed to delete the previous file {token} {response}", acctok.Short(), response.Status);
            }
            else
            {
                log.LogInformation("Previous file unlocked and removed successfully {token}", acctok.Short());
            }
        }

        log.LogInformation("Final save completed {filename} {token}", newName, acctok.Short());
        return Reply("File saved successfully", StatusCodes.Status200OK);
    }
}
