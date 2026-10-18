/// <summary>
/// The CodiMD-specific code used by the WOPI bridge.
/// </summary>
partial class CodiMD(BridgeHttp http, WopiClient wopi, ILogger log) :
    BridgeApp(http, wopi, log)
{
    const string tooLarge = "File is too large to be edited in CodiMD. Please reduce its size with a regular text editor and try again.";

    /// <summary>
    /// Uploads have links like '/uploads/upload_542a360ddefe1e21ad1b8c85207d9365.*'.
    /// </summary>
    [GeneratedRegex(@"\/uploads\/upload_\w{32}\.\w+")]
    private static partial Regex UploadRegex();

    /// <summary>
    /// CodiMD integrates Prometheus metrics, let's probe if they exist.
    /// </summary>
    protected override async Task Probe()
    {
        try
        {
            using var request = new HttpRequestMessage(HttpMethod.Head, $"{AppUrl}/metrics/codimd");
            using var response = await Http.NoRedirect.SendAsync(request);
            if (response.StatusCode != HttpStatusCode.OK)
            {
                Log.LogError("The provided URL does not seem to be a CodiMD instance {appurl}", AppUrl);
                throw new AppFailure();
            }

            Log.LogInformation("Successfully connected to CodiMD {appurl}", AppUrl);
        }
        catch (Exception exception) when (IsConnectionError(exception))
        {
            Log.LogError("Exception raised attempting to connect to CodiMD {exception}", exception.Message);
            throw new AppFailure("", exception);
        }
    }

    public override Task<string> GetRedirectUrl(ViewMode viewMode, string wopiSrc, string acctok, string docId, string filename, string displayName, string? revaToken)
    {
        if (!viewMode.CanWrite())
        {
            // read-only mode: use the publish view of CodiMD
            return Task.FromResult($"{AppExtUrl}/{docId}/publish");
        }

        var mode = viewMode == ViewMode.Preview ? "view" : "both";
        var parameters = new List<KeyValuePair<string, string>>
        {
            new("wopiSrc", wopiSrc),
            new("accessToken", acctok),
            new("disableEmbedding", PosixPath.Extension(filename) != ".zmd" ? "true" : "false"),
            new("displayName", displayName),
            new("path", PosixPath.DirName(filename))
        };
        if (!string.IsNullOrEmpty(revaToken))
        {
            parameters.Add(new("revaToken", revaToken));
        }

        return Task.FromResult($"{AppExtUrl}/{docId}?{mode}&{UrlCodec.UrlEncode(parameters)}");
    }

    // Cloud storage to CodiMD

    /// <summary>
    /// Unzips the given input buffer uploading the content to CodiMD, and returns the contained .md file.
    /// </summary>
    async Task<byte[]> UnzipAttachments(byte[] input, string? docId, string wopiSrc, string acctok)
    {
        byte[]? mdDoc = null;
        var renamed = new List<(string From, string To)>();
        try
        {
            using var zip = new ZipArchive(new MemoryStream(input), ZipArchiveMode.Read);
            foreach (var entry in zip.Entries)
            {
                var name = entry.FullName;
                Log.LogDebug("Extracting attachment {name}", name);
                if (PosixPath.Extension(name) == ".md")
                {
                    mdDoc = await ReadEntry(entry);
                    continue;
                }

                // first check if the file already exists in CodiMD
                using (var head = new HttpRequestMessage(HttpMethod.Head, $"{AppUrl}/uploads/{name}"))
                using (var response = await Http.NoRedirect.SendAsync(head))
                {
                    if (response.StatusCode == HttpStatusCode.OK)
                    {
                        if (response.Content.Headers.ContentLength == entry.Length)
                        {
                            // yes (assume that hashed filename AND size matching is a good enough content match!)
                            Log.LogDebug("Skipped existing attachment {filename}", name);
                            continue;
                        }

                        // we have a collision
                        Log.LogWarning("Attachment collision detected {filename}", name);
                        // append a random letter to the filename, and replace its reference in the document
                        var (root, extension) = PosixPath.SplitExt(name);
                        var newName = $"{root}_{(char) Random.Shared.Next('A', 'Z' + 1)}{extension}";
                        renamed.Add((name, newName));
                        name = newName;
                    }
                }

                // OK, let's upload
                Log.LogDebug("Pushing attachment {filename}", name);
                var query = Query(("generateFilename", "false"), ("WOPISrc", wopiSrc), ("docId", docId ?? ""), ("accessToken", acctok));
                using var form = new MultipartFormDataContent();
                form.Add(new ByteArrayContent(await ReadEntry(entry)), "image", name);
                using var upload = await Http.Follow.PostAsync($"{AppUrl}/uploadimage?{query}", form);
                if (upload.StatusCode != HttpStatusCode.OK)
                {
                    Log.LogError("Failed to push included file {filename} {httpcode}", name, (int) upload.StatusCode);
                }
            }
        }
        catch (InvalidDataException exception)
        {
            throw new AppFailure("The file is not in the expected zipped format", exception);
        }
        catch (Exception exception) when (IsConnectionError(exception))
        {
            Log.LogError("Exception raised attempting to connect to CodiMD {exception}", exception.Message);
            throw new AppFailure("Failed to connect to CodiMD", exception);
        }

        if (mdDoc == null)
        {
            return [];
        }

        var text = Body(mdDoc);
        foreach (var (from, to) in renamed)
        {
            text = text.Replace(from, to);
        }

        // for backwards compatibility, drop the hardcoded reverse proxy paths if found in the document
        return Encoding.UTF8.GetBytes(text.Replace("/byoa/codimd/", "/"));
    }

    static async Task<byte[]> ReadEntry(ZipArchiveEntry entry)
    {
        await using var stream = await entry.OpenAsync();
        using var memory = new MemoryStream();
        await stream.CopyToAsync(memory);
        return memory.ToArray();
    }

    /// <summary>
    /// Fetches a given document from CodiMD, throws <see cref="AppFailure"/> in case of errors.
    /// </summary>
    public override async Task<byte[]> Fetch(BridgeLock wopiLock, string acctok)
    {
        try
        {
            var separator = wopiLock.Doc.StartsWith('/') ? "" : "/";
            using var response = await Http.Follow.GetAsync($"{AppUrl}{separator}{wopiLock.Doc}/download");
            var content = await response.Content.ReadAsByteArrayAsync();
            if (response.StatusCode != HttpStatusCode.OK)
            {
                Log.LogError("Unable to fetch document from CodiMD {token} {response}", acctok.Short(), $"{(int) response.StatusCode}: {Body(content, 50)}");
                throw new AppFailure();
            }

            return content;
        }
        catch (Exception exception) when (IsConnectionError(exception))
        {
            Log.LogError("Exception raised attempting to connect to CodiMD {exception}", exception.Message);
            throw new AppFailure("Failed to connect to CodiMD", exception);
        }
    }

    /// <summary>
    /// Copies the document from storage to CodiMD.
    /// </summary>
    public override async Task<BridgeLock> LoadFromStorage(JsonNode fileMetadata, string wopiSrc, string acctok, string? docId)
    {
        // WOPI GetFile
        var response = await Wopi.Request(wopiSrc, acctok, HttpMethod.Get, contents: []);
        if (response.Status != StatusCodes.Status200OK)
        {
            throw new AppFailure($"Unable to fetch file from storage, got HTTP {response.Status}");
        }

        var mdFile = response.Content;
        var wasBundle = PosixPath.Extension((string) fileMetadata["BaseFileName"]!) == ".zmd";

        // if it's a bundled file, unzip it and push the attachments in the appropriate folder
        var mdDoc = mdFile;
        if (wasBundle &&
            mdFile.Length > 0)
        {
            mdDoc = await UnzipAttachments(mdFile, docId, wopiSrc, acctok);
        }

        string text;
        try
        {
            text = new UTF8Encoding(false, true).GetString(mdDoc);
        }
        catch (DecoderFallbackException exception)
        {
            Log.LogWarning("Invalid UTF-8 content found in file {exception}", exception.Message);
            throw new AppFailure("File contains an invalid UTF-8 character, was it corrupted? Please fix it in a regular editor before opening it in CodiMD.", exception);
        }

        // if the file was created on Windows, convert \r\n to \n for CodiMD to correctly edit it
        if (text.Contains("\r\n"))
        {
            text = text.Replace("\r\n", "\n");
            mdDoc = Encoding.UTF8.GetBytes(text);
        }

        try
        {
            if (docId == null)
            {
                docId = await PushReadOnly(mdDoc, acctok);
            }
            else
            {
                docId = await Push(docId, text, acctok);
            }
        }
        catch (Exception exception) when (IsConnectionError(exception))
        {
            Log.LogError("Exception raised attempting to connect to CodiMD {exception}", exception.Message);
            throw new AppFailure("", exception);
        }

        // generate and return a WOPI lock structure for this document
        return WopiClient.GenerateLock(docId, fileMetadata, mdDoc, acctok, false);
    }

    /// <summary>
    /// Read-only case: push the doc to a newly generated note with a random docid.
    /// </summary>
    async Task<string> PushReadOnly(byte[] mdDoc, string acctok)
    {
        using var request = new HttpRequestMessage(HttpMethod.Post, $"{AppUrl}/new?mode=locked");
        request.Content = new ByteArrayContent(mdDoc);
        request.Content.Headers.ContentType = new("text/markdown");
        request.Headers.Authorization = new("Bearer", ApiKey);
        using var response = await Http.NoRedirect.SendAsync(request);
        if (response.StatusCode == HttpStatusCode.RequestEntityTooLarge)
        {
            Log.LogError("File is too large to be edited in CodiMD {token}", acctok.Short());
            throw new AppFailure(tooLarge);
        }

        if (response.StatusCode != HttpStatusCode.Found)
        {
            Log.LogError("Unable to push read-only document to CodiMD {token} {response}", acctok.Short(), (int) response.StatusCode);
            throw new AppFailure();
        }

        var docId = LocationId(response);
        Log.LogInformation("Pushed read-only document to CodiMD {docid} {token}", docId, acctok.Short());
        return docId;
    }

    async Task<string> Push(string docId, string text, string acctok)
    {
        // reserve the given docid in CodiMD via a HEAD request
        using (var request = new HttpRequestMessage(HttpMethod.Head, $"{AppUrl}/{docId}"))
        {
            request.Headers.Authorization = new("Bearer", ApiKey);
            using var response = await Http.NoRedirect.SendAsync(request);
            if (response.StatusCode is not (HttpStatusCode.OK or HttpStatusCode.Found))
            {
                Log.LogError("Unable to reserve note hash in CodiMD {token} {response}", acctok.Short(), (int) response.StatusCode);
                throw new AppFailure();
            }

            // check if the target docid is real or is a redirect
            if (response.StatusCode == HttpStatusCode.Found)
            {
                var newDocId = LocationId(response);
                Log.LogInformation("Document got aliased in CodiMD {olddocid} {docid} {token}", docId, newDocId, acctok.Short());
                docId = newDocId;
            }
        }

        // push the document to CodiMD with the update API
        using (var request = new HttpRequestMessage(HttpMethod.Put, $"{AppUrl}/api/notes/{docId}"))
        {
            request.Headers.Authorization = new("Bearer", ApiKey);
            request.Content = ContentOf(
                new JsonObject
                {
                    ["content"] = text
                }.ToJsonString(),
                "application/json");
            using var response = await Http.Follow.SendAsync(request);
            if (response.StatusCode == HttpStatusCode.Forbidden)
            {
                // the file got unlocked because of no activity, yet some user is there: let it go
                Log.LogWarning("Document was being edited in CodiMD, redirecting user {token}", acctok.Short());
            }
            else if (response.StatusCode == HttpStatusCode.RequestEntityTooLarge)
            {
                Log.LogError("File is too large to be edited in CodiMD {docid} {token}", docId, acctok.Short());
                throw new AppFailure(tooLarge);
            }
            else if (response.StatusCode != HttpStatusCode.OK)
            {
                Log.LogError("Unable to push document to CodiMD {docid} {token} {response}", docId, acctok.Short(), (int) response.StatusCode);
                throw new AppFailure();
            }
        }

        Log.LogInformation("Pushed document to CodiMD {docid} {token}", docId, acctok.Short());
        return docId;
    }

    // CodiMD to cloud storage

    /// <summary>
    /// Parses a markdown file and generates a zip file containing all included files.
    /// </summary>
    async Task<(byte[]? Bundle, BridgeReply? Reply)> GetAttachments(string mdDoc, string docFileName, bool forceZip)
    {
        BridgeReply? reply = null;
        using var buffer = new MemoryStream();
        var count = 0;
        using (var zip = new ZipArchive(buffer, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (Match match in UploadRegex().Matches(mdDoc))
            {
                var attachment = match.Value;
                Log.LogDebug("Fetching attachment {url}", attachment);
                byte[] content;
                try
                {
                    using var response = await Http.Follow.GetAsync($"{AppUrl}{attachment}");
                    if (response.StatusCode != HttpStatusCode.OK)
                    {
                        throw new HttpRequestException(((int) response.StatusCode).ToString(CultureInfo.InvariantCulture));
                    }

                    content = await response.Content.ReadAsByteArrayAsync();
                }
                catch (Exception exception) when (IsConnectionError(exception))
                {
                    Log.LogError("Failed to fetch included file, skipping {path} {type} {error}", attachment, exception.GetType().Name, exception.Message);
                    // also notify the user
                    reply = WopiClient.Reply("Failed to include a referenced picture in the saved file", StatusCodes.Status404NotFound);
                    continue;
                }

                await AddEntry(zip, attachment.Split('/')[^1], content);
                count++;
            }

            if (!forceZip &&
                count == 0)
            {
                // no attachments actually found
                return (null, reply);
            }

            // also include the markdown file itself
            await AddEntry(zip, docFileName, Encoding.UTF8.GetBytes(mdDoc));
        }

        return (buffer.ToArray(), reply);
    }

    static async Task AddEntry(ZipArchive zip, string name, byte[] content)
    {
        var entry = zip.CreateEntry(name, CompressionLevel.NoCompression);
        await using var stream = await entry.OpenAsync();
        await stream.WriteAsync(content);
    }

    /// <summary>
    /// Copies the document from CodiMD back to storage.
    /// </summary>
    public override async Task<BridgeReply> SaveToStorage(string wopiSrc, string acctok, bool isClose, BridgeLock wopiLock)
    {
        // get document from CodiMD
        byte[] mdDoc;
        try
        {
            Log.LogInformation("Fetching file from CodiMD {isclose} {appurl} {token}", isClose, $"{AppUrl}{wopiLock.Doc}", acctok.Short());
            mdDoc = await Fetch(wopiLock, acctok);
        }
        catch (AppFailure)
        {
            // return a non-fatal error
            return WopiClient.Reply("File not saved, error in fetching document from CodiMD. Will try again later", StatusCodes.Status424FailedDependency);
        }

        if (isClose &&
            Wopi.CheckForNoChanges(mdDoc, wopiLock, acctok))
        {
            return new("{}", StatusCodes.Status202Accepted);
        }

        // check if we have attachments
        var wasBundle = PosixPath.Extension(wopiLock.FileName) == ".zmd";
        byte[]? bundle = null;
        BridgeReply? attachmentsReply = null;
        // in disablezip mode, preserve existing .zmd files but don't create new ones
        if (!DisableZip ||
            wasBundle)
        {
            (bundle, attachmentsReply) = await GetAttachments(Body(mdDoc), wopiLock.FileName.Replace(".zmd", ".md"), wasBundle);
        }

        // WOPI PutFile for the file or the bundle if it already existed
        if (wasBundle ^ (bundle == null) ||
            !isClose)
        {
            var response = await Wopi.Request(
                wopiSrc,
                acctok,
                HttpMethod.Post,
                wasBundle ? bundle : mdDoc,
                new()
                {
                    ["X-WOPI-Lock"] = wopiLock.ToJson()
                });
            var reply = Wopi.HandlePutFile("PutFile", wopiSrc, response);
            if (reply != null)
            {
                return reply;
            }

            var lockReply = await Wopi.RefreshDigestAndLock(wopiSrc, acctok, wopiLock, mdDoc);
            // combine the responses
            return attachmentsReply ?? lockReply;
        }

        // on close, use saveas for either the new bundle, if this is the first time we have attachments,
        // or the single md file, if there are no more attachments.
        var target = $"{PosixPath.SplitExt(wopiLock.FileName).Root}{(bundle != null ? ".zmd" : ".md")}";
        return await Wopi.SaveAs(wopiSrc, acctok, wopiLock, target, bundle ?? mdDoc);
    }
}
