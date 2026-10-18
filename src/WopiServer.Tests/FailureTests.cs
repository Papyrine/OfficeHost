/// <summary>
/// The handling of storage failures, desktop locks and fallbacks, through a <see cref="FaultyStorage"/>.
/// </summary>
public class FailureTests
{
    const string databuf = "file content";

    static StorageException Failure(string message = "Input/output error") =>
        new(message);

    static string? Header(HttpResponseMessage response, string name)
    {
        if (response.Headers.TryGetValues(name, out var values))
        {
            return string.Join(", ", values);
        }

        return null;
    }

    sealed class Context(TestServerHost host, FaultyStorage storage, string fileId, string token) : IAsyncDisposable
    {
        public TestServerHost Host { get; } = host;
        public FaultyStorage Storage { get; } = storage;
        public string FileId { get; } = fileId;
        public string Token { get; } = token;

        /// <summary>
        /// Fails the given operation on the paths ending with the suffix.
        /// </summary>
        public void Fail(string operation, string suffix, Exception exception) =>
            Storage.Fault = (op, path) => op == operation && path.EndsWith(suffix, StringComparison.Ordinal) ? exception : null;

        public Task<HttpResponseMessage> Post(string operation, params (string Name, string Value)[] headers) =>
            Host.Post(FileId, Token, operation, headers);

        public ValueTask DisposeAsync() =>
            Host.DisposeAsync();
    }

    static async Task<Context> Start(string file = "/test.docx", string content = databuf, string extraConfig = "")
    {
        FaultyStorage? storage = null;
        var host = await TestServerHost.Start(extraConfig, wrapStorage: _ => storage = new(_));
        Directory.CreateDirectory(Path.GetDirectoryName(host.PathOf(file))!);
        await File.WriteAllTextAsync(host.PathOf(file), content);
        var (fileId, token, _) = await host.Open(file);
        return new(host, storage!, fileId, token);
    }

    // Lock

    /// <summary>
    /// The save time is only used to detect external modifications, so failing to store it does not fail the lock.
    /// </summary>
    [Test]
    public async Task Lock_SaveTimeFails()
    {
        await using var context = await Start();
        context.Fail(nameof(IStorage.SetXattr), "/test.docx", Failure());
        using var response = await context.Post("LOCK", ("X-WOPI-Lock", "L"));
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.OK);
    }

    [Test]
    public async Task Lock_StorageFails()
    {
        await using var context = await Start();
        context.Fail(nameof(IStorage.SetLock), "/test.docx", Failure());
        using var response = await context.Post("LOCK", ("X-WOPI-Lock", "L"));
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.InternalServerError);
        await Assert.That(Header(response, "X-WOPI-ServerError")).IsEqualTo("I/O Error, please contact support");
    }

    /// <summary>
    /// Taking the lock already held refreshes it: a failing refresh is a server error.
    /// </summary>
    [Test]
    public async Task Lock_RefreshFails()
    {
        await using var context = await Start();
        using (var first = await context.Post("LOCK", ("X-WOPI-Lock", "L")))
        {
            await Assert.That(first.StatusCode).IsEqualTo(HttpStatusCode.OK);
        }

        context.Fail(nameof(IStorage.RefreshLock), "/test.docx", Failure());
        using var response = await context.Post("LOCK", ("X-WOPI-Lock", "L"));
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.InternalServerError);
    }

    [Test]
    public async Task Lock_UnknownFileStat()
    {
        await using var context = await Start();
        context.Fail(nameof(IStorage.StatX), "/test.docx", Failure());
        using var response = await context.Post("LOCK", ("X-WOPI-Lock", "L"));
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.InternalServerError);
    }

    /// <summary>
    /// The LibreOffice-compatible lock file is only for interoperability: failing to write it does not fail the lock.
    /// </summary>
    [Test]
    public async Task Lock_LibreOfficeLockFileFails()
    {
        await using var context = await Start("/test.odt");
        context.Fail(nameof(IStorage.WriteFile), "#", new StorageException(Common.AccessError));
        using var response = await context.Post("LOCK", ("X-WOPI-Lock", "L"));
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.OK);
        await Assert.That(File.Exists(context.Host.PathOf("/.~lock.test.odt#"))).IsFalse();
    }

    /// <summary>
    /// A LibreOffice-compatible lock file the server wrote itself (here left over as its removal failed) does not
    /// block a new lock. The Python implementation looked for "WOPIServer" in a file holding "wopiserver".
    /// </summary>
    [Test]
    public async Task Lock_OwnLibreOfficeLockFileLeftOver()
    {
        await using var context = await Start("/test.odt");
        await File.WriteAllTextAsync(context.Host.PathOf("/.~lock.test.odt#"), $",root via web collaboration,cernbox,{DateTime.Now:dd.MM.yyyy HH:mm},wopiserver;");
        context.Fail(nameof(IStorage.RemoveFile), "#", Failure());
        using var response = await context.Post("LOCK", ("X-WOPI-Lock", "L"));
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.OK);
    }

    /// <summary>
    /// A LibreOffice lock file older than the lock expiration is stale, and is taken over.
    /// </summary>
    [Test]
    public async Task Lock_StaleLibreOfficeLockFile()
    {
        await using var context = await Start("/test.odt");
        await File.WriteAllTextAsync(context.Host.PathOf("/.~lock.test.odt#"), $",root via web collaboration,cernbox,{DateTime.Now.AddHours(-2):dd.MM.yyyy HH:mm},wopiserver;");
        context.Fail(nameof(IStorage.RemoveFile), "#", Failure());
        using var response = await context.Post("LOCK", ("X-WOPI-Lock", "L"));
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.OK);
    }

    /// <summary>
    /// A lock set on the storage by something else than this server is an external lock.
    /// </summary>
    [Test]
    public async Task GetLock_ForeignLock()
    {
        await using var context = await Start();
        await context.Storage.SetLock("default", "/test.docx", "0:0", "another app", "not a WOPI lock");
        using var response = await context.Post("GET_LOCK");
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.OK);
        await Assert.That(Header(response, "X-WOPI-Lock")).IsEqualTo(WopiUtils.ExternalLock);
        using var open = await context.Host.OpenInApp("/test.docx");
        var reason = (string?) JsonNode.Parse(await open.Content.ReadAsStringAsync())!["forced-viewmode-reason"];
        await Assert.That(reason).IsEqualTo("This file is opened exclusively by another app or user and cannot be edited");
    }

    // Unlock

    [Test]
    public async Task Unlock_StorageFails()
    {
        await using var context = await Start();
        using (var setLock = await context.Post("LOCK", ("X-WOPI-Lock", "L")))
        {
            await Assert.That(setLock.StatusCode).IsEqualTo(HttpStatusCode.OK);
        }

        context.Fail(nameof(IStorage.Unlock), "/test.docx", Failure());
        using var response = await context.Post("UNLOCK", ("X-WOPI-Lock", "L"));
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.InternalServerError);
    }

    /// <summary>
    /// The removal of the LibreOffice-compatible lock file on unlock is best effort.
    /// </summary>
    [Test]
    public async Task Unlock_LibreOfficeLockFileRemovalFails()
    {
        await using var context = await Start("/test.odt");
        using (var setLock = await context.Post("LOCK", ("X-WOPI-Lock", "L")))
        {
            await Assert.That(setLock.StatusCode).IsEqualTo(HttpStatusCode.OK);
        }

        context.Fail(nameof(IStorage.RemoveFile), "#", Failure());
        using var response = await context.Post("UNLOCK", ("X-WOPI-Lock", "L"));
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.OK);
    }

    // PutFile

    /// <summary>
    /// A failed save is a server error, and the content is kept in the recovery space.
    /// </summary>
    [Test]
    public async Task PutFile_WriteFails()
    {
        await using var context = await Start();
        using (var setLock = await context.Post("LOCK", ("X-WOPI-Lock", "L")))
        {
            await Assert.That(setLock.StatusCode).IsEqualTo(HttpStatusCode.OK);
        }

        context.Fail(nameof(IStorage.WriteFile), "/test.docx", Failure());
        using var response = await context.Host.PutFile(context.FileId, context.Token, "precious content", ("X-WOPI-Lock", "L"));
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.InternalServerError);
        var recovered = Directory.GetFiles(context.Host.RecoveryPath).Single();
        await Assert.That(await File.ReadAllTextAsync(recovered)).IsEqualTo("precious content");
    }

    /// <summary>
    /// The editnew case fails the same way.
    /// </summary>
    [Test]
    public async Task PutFile_EditNewWriteFails()
    {
        await using var context = await Start(content: "");
        context.Fail(nameof(IStorage.WriteFile), "/test.docx", Failure());
        using var response = await context.Host.PutFile(context.FileId, context.Token, "first content");
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.InternalServerError);
        await Assert.That(await File.ReadAllTextAsync(Directory.GetFiles(context.Host.RecoveryPath).Single())).IsEqualTo("first content");
    }

    /// <summary>
    /// A desktop lock taken while the file is locked online: the save is refused, and kept as a conflict copy.
    /// </summary>
    [Test]
    public async Task PutFile_ExternalLock()
    {
        await using var context = await Start();
        using (var setLock = await context.Post("LOCK", ("X-WOPI-Lock", "L")))
        {
            await Assert.That(setLock.StatusCode).IsEqualTo(HttpStatusCode.OK);
        }

        await File.WriteAllTextAsync(context.Host.PathOf("/~$test.docx"), "lock");
        using var response = await context.Host.PutFile(context.FileId, context.Token, "web content", ("X-WOPI-Lock", "L"));
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.Conflict);
        await Assert.That(Header(response, "X-WOPI-LockFailureReason")).IsEqualTo("Cannot overwrite file edited by 0:0 via Microsoft Office for Desktop, conflict copy created");
        var copy = Directory.GetFiles(context.Host.StoragePath, "test-webconflict-*.docx").Single();
        await Assert.That(await File.ReadAllTextAsync(copy)).IsEqualTo("web content");
    }

    /// <summary>
    /// When the folder of the file is not writable, the conflict copy goes to the user's home; when that fails too, to
    /// the recovery space.
    /// </summary>
    [Test]
    [Arguments(true)]
    [Arguments(false)]
    public async Task PutFile_ConflictCopyFallbacks(bool homeWritable)
    {
        await using var context = await Start("/shared/test.docx");
        Directory.CreateDirectory(context.Host.PathOf("/home/root"));
        using (var setLock = await context.Post("LOCK", ("X-WOPI-Lock", "L")))
        {
            await Assert.That(setLock.StatusCode).IsEqualTo(HttpStatusCode.OK);
        }

        context.Storage.Fault = (operation, path) =>
            operation == nameof(IStorage.WriteFile) && path.Contains("webconflict") && (path.StartsWith("/shared/") || !homeWritable) ? new StorageException(Common.AccessError) : null;
        File.SetLastWriteTimeUtc(context.Host.PathOf("/shared/test.docx"), DateTime.UtcNow.AddMinutes(5));
        using var response = await context.Host.PutFile(context.FileId, context.Token, "web content", ("X-WOPI-Lock", "L"));
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.Conflict);
        if (homeWritable)
        {
            await Assert.That(Header(response, "X-WOPI-LockFailureReason")).EndsWith(", conflict copy created");
            await Assert.That(Directory.GetFiles(context.Host.PathOf("/home/root"), "test-webconflict-*.docx")).Count().IsEqualTo(1);
        }
        else
        {
            await Assert.That(Header(response, "X-WOPI-LockFailureReason")).EndsWith(", please contact support to recover it");
            await Assert.That(await File.ReadAllTextAsync(Directory.GetFiles(context.Host.RecoveryPath).Single())).IsEqualTo("web content");
        }
    }

    // PutRelativeFile

    /// <summary>
    /// When the folder of the file is not writable (a single-file share), "save as" goes to the user's home.
    /// </summary>
    [Test]
    [Arguments(true, HttpStatusCode.OK)]
    [Arguments(false, HttpStatusCode.NotImplemented)]
    public async Task PutRelativeFile_HomeFallback(bool homeWritable, HttpStatusCode expected)
    {
        await using var context = await Start("/shared/test.docx");
        Directory.CreateDirectory(context.Host.PathOf("/home/root"));
        context.Storage.Fault = (operation, path) =>
            operation == nameof(IStorage.WriteFile) && (path.StartsWith("/shared/") || !homeWritable) ? new StorageException(Common.AccessError) : null;
        using var response = await context.Post("PUT_RELATIVE", ("X-WOPI-SuggestedTarget", ".pdf"));
        await Assert.That(response.StatusCode).IsEqualTo(expected);
        await Assert.That(File.Exists(context.Host.PathOf("/home/root/test.pdf"))).IsEqualTo(homeWritable);
    }

    [Test]
    public async Task PutRelativeFile_WriteFails()
    {
        await using var context = await Start();
        context.Fail(nameof(IStorage.WriteFile), ".pdf", Failure());
        using var response = await context.Post("PUT_RELATIVE", ("X-WOPI-SuggestedTarget", ".pdf"));
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.InternalServerError);
    }

    /// <summary>
    /// Finding a free name for a suggested target needs the storage to tell a missing file from other failures.
    /// </summary>
    [Test]
    public async Task PutRelativeFile_StatFails()
    {
        await using var context = await Start();
        context.Fail(nameof(IStorage.Stat), ".pdf", Failure());
        using var response = await context.Post("PUT_RELATIVE", ("X-WOPI-SuggestedTarget", ".pdf"));
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.InternalServerError);
    }

    // Other operations

    [Test]
    public async Task DeleteFile_StorageFails()
    {
        await using var context = await Start();
        context.Fail(nameof(IStorage.RemoveFile), "/test.docx", Failure());
        using var response = await context.Post("DELETE");
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.InternalServerError);
    }

    [Test]
    public async Task RenameFile_StorageFails()
    {
        await using var context = await Start(extraConfig: "enablerename = True");
        context.Fail(nameof(IStorage.RenameFile), "/test.docx", Failure());
        using var response = await context.Post("RENAME_FILE", ("X-WOPI-Lock", ""), ("X-WOPI-RequestedName", "renamed"));
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.InternalServerError);
        await Assert.That(Header(response, "X-WOPI-InvalidFileNameError")).IsEqualTo("Failed to rename: Input/output error");
    }

    [Test]
    public async Task RenameFile_LockedByStorage()
    {
        await using var context = await Start(extraConfig: "enablerename = True");
        context.Fail(nameof(IStorage.RenameFile), "/test.docx", new StorageException(Common.ExclError));
        using var response = await context.Post("RENAME_FILE", ("X-WOPI-Lock", ""), ("X-WOPI-RequestedName", "renamed"));
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.NotImplemented);
        await Assert.That(Header(response, "X-WOPI-InvalidFileNameError")).IsEqualTo("Cannot rename/move unlocked file");
    }

    /// <summary>
    /// Renaming an OpenDocument file also renames its LibreOffice-compatible lock file.
    /// </summary>
    [Test]
    public async Task RenameFile_MovesLibreOfficeLockFile()
    {
        await using var context = await Start("/test.odt", extraConfig: "enablerename = True");
        using (var setLock = await context.Post("LOCK", ("X-WOPI-Lock", "L")))
        {
            await Assert.That(setLock.StatusCode).IsEqualTo(HttpStatusCode.OK);
        }

        using var response = await context.Post("RENAME_FILE", ("X-WOPI-Lock", "L"), ("X-WOPI-RequestedName", "renamed"));
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.OK);
        await Assert.That(File.Exists(context.Host.PathOf("/.~lock.renamed.odt#"))).IsTrue();
        await Assert.That(File.Exists(context.Host.PathOf("/.~lock.test.odt#"))).IsFalse();
    }

    [Test]
    public async Task PutUserInfo_StorageFails()
    {
        await using var context = await Start();
        context.Fail(nameof(IStorage.SetXattr), "/test.docx", Failure());
        using var request = new HttpRequestMessage(HttpMethod.Post, $"/wopi/files/{context.FileId}?access_token={context.Token}");
        request.Headers.Add("X-WOPI-Override", "PUT_USER_INFO");
        request.Content = new StringContent("info");
        using var response = await context.Host.Client.SendAsync(request);
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.InternalServerError);
    }

    [Test]
    public async Task GetFile_ReadFails()
    {
        await using var context = await Start();
        context.Fail(nameof(IStorage.ReadFile), "/test.docx", Failure());
        using var response = await context.Host.Client.GetAsync($"/wopi/files/{context.FileId}/contents?access_token={context.Token}");
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.InternalServerError);
    }

    /// <summary>
    /// As in the Python implementation, any failure to stat the file is reported as an unknown file.
    /// </summary>
    [Test]
    public async Task CheckFileInfo_StatFails()
    {
        await using var context = await Start();
        context.Fail(nameof(IStorage.StatX), "/test.docx", Failure());
        using var response = await context.Host.Client.GetAsync($"/wopi/files/{context.FileId}?access_token={context.Token}");
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.NotFound);
    }

    /// <summary>
    /// A storage that throws something else than a StorageException is an unexpected server error.
    /// </summary>
    [Test]
    public async Task UnexpectedException()
    {
        await using var context = await Start();
        context.Fail(nameof(IStorage.StatX), "/test.docx", new InvalidOperationException("bug"));
        using var response = await context.Host.Client.GetAsync($"/wopi/files/{context.FileId}?access_token={context.Token}");
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.InternalServerError);
        await Assert.That(await response.Content.ReadAsStringAsync()).IsEqualTo("Internal error, please contact support");
        await Assert.That(Header(response, "X-WOPI-ServerError")).IsEqualTo("Server error");
        await Assert.That(context.Host.Log.ToString()).Contains("Unexpected exception caught");
    }

    // CheckFileInfo details

    /// <summary>
    /// The CSPP+ SequenceNumber comes from the version part of EOS etags (`version:checksum`).
    /// </summary>
    [Test]
    [Arguments("12345:7fd2729c", 12345L)]
    [Arguments("notanumber:7fd2729c", null)]
    [Arguments("1700000000.1234567", 17000000001234567L)]
    [Arguments("W/abc", null)]
    public async Task CheckFileInfo_SequenceNumber(string etag, long? expected)
    {
        await using var context = await Start("/test.wopitest");
        context.Storage.MapStat = _ => _ with { Etag = etag };
        using var response = await context.Host.Client.GetAsync($"/wopi/files/{context.FileId}?access_token={context.Token}");
        var info = JsonNode.Parse(await response.Content.ReadAsStringAsync())!;
        await Assert.That((long?) info["SequenceNumber"]).IsEqualTo(expected);
        await Assert.That((string?) info["Version"]).IsEqualTo($"v{etag}");
    }

    /// <summary>
    /// Out of Debug, the URLs carrying the access token are redacted from the log.
    /// </summary>
    [Test]
    [Arguments("Info", true)]
    [Arguments("Debug", false)]
    public async Task CheckFileInfo_LogRedaction(string logLevel, bool redacted)
    {
        await using var context = await Start(extraConfig: $"loglevel = {logLevel}\ndownloadurl = https://wopi/download");
        using var response = await context.Host.Client.GetAsync($"/wopi/files/{context.FileId}?access_token={context.Token}");
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.OK);
        var line = context.Host.Log.ToString().Split('\n').Single(_ => _.Contains("\"File metadata response\""));
        await Assert.That(line.Contains("_redacted_")).IsEqualTo(redacted);
        await Assert.That(line.Contains("https://wopi/download?access_token=")).IsEqualTo(!redacted);
    }
}
