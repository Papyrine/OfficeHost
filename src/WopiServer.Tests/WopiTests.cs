/// <summary>
/// The WOPI protocol, end to end over the local storage.
/// </summary>
public class WopiTests
{
    static async Task<string> Body(HttpResponseMessage response) =>
        await response.Content.ReadAsStringAsync();

    static string? Header(HttpResponseMessage response, string name)
    {
        if (response.Headers.TryGetValues(name, out var values))
        {
            return string.Join(", ", values);
        }

        return null;
    }

    [Test]
    public async Task Index()
    {
        await using var host = await TestServerHost.Start();
        using var response = await host.Client.GetAsync("/wopi");
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.OK);
        var body = await Body(response);
        await Assert.That(body).Contains("Storage type: <span style=\"font-family:monospace\">local</span>");
        await Assert.That(body).Contains("Health status: <span style=\"font-family:monospace\">Warning</span>");
    }

    [Test]
    public async Task OpenInApp_Unauthorized()
    {
        await using var host = await TestServerHost.Start();
        using var response = await host.Client.GetAsync("/wopi/iop/openinapp?fileid=/test.docx");
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.Unauthorized);
    }

    [Test]
    public async Task OpenInApp_InvalidArguments()
    {
        await using var host = await TestServerHost.Start();
        using var invalidViewMode = await host.OpenInApp("/test.docx", "VIEW_MODE_WHATEVER");
        await Assert.That(invalidViewMode.StatusCode).IsEqualTo(HttpStatusCode.BadRequest);
        await Assert.That(await Body(invalidViewMode)).IsEqualTo("Missing or invalid viewmode argument");
        using var missingFile = await host.OpenInApp("/missing.docx");
        await Assert.That(missingFile.StatusCode).IsEqualTo(HttpStatusCode.NotFound);
        using var folder = await host.OpenInApp("/");
        await Assert.That(folder.StatusCode).IsEqualTo(HttpStatusCode.NotFound);
    }

    [Test]
    public async Task OpenInApp_ForcesReadOnly_ForLegacyFormats()
    {
        await using var host = await TestServerHost.Start();
        await File.WriteAllTextAsync(host.PathOf("/legacy.doc"), "content");
        var (_, _, response) = await host.Open("/legacy.doc");
        await Assert.That((string?) response["app-url"]).StartsWith("https://office/view?WOPISrc=");
        await Assert.That((string?) response["forced-viewmode-reason"]).IsEqualTo("App does not support edit for this file type");
    }

    [Test]
    public async Task OpenInApp_ForcesReadOnly_WithMicrosoftOfficeLock()
    {
        await using var host = await TestServerHost.Start();
        await File.WriteAllTextAsync(host.PathOf("/test.docx"), "content");
        await File.WriteAllTextAsync(host.PathOf("/~$test.docx"), "lock");
        var (_, _, response) = await host.Open("/test.docx");
        await Assert.That((string?) response["forced-viewmode-reason"]).IsEqualTo("This file is opened exclusively by 0:0 via Microsoft Office for Desktop and cannot be edited");
    }

    [Test]
    public async Task CheckFileInfo()
    {
        await using var host = await TestServerHost.Start(
            """
            brandingname = CERNBox
            downloadurl = https://wopi/download
            hostediturl = https://efss/external-<app><path>
            """);
        await File.WriteAllTextAsync(host.PathOf("/test.docx"), "content");
        var (fileId, token, open) = await host.Open("/test.docx");
        await Assert.That((string?) open["app-url"]).StartsWith("https://office/edit?WOPISrc=http%3A%2F%2Flocalhost%2Fwopi%2Ffiles%2Flocal%21");
        using var response = await host.Client.GetAsync($"/wopi/files/{fileId}?access_token={token}");
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.OK);
        await VerifyJson(await Body(response))
            .ScrubMembers("ServerTime", "AccessTokenExpiry", "LastModifiedTime", "Version", "HostViewUrl", "DownloadUrl", "FileUrl");
    }

    [Test]
    public async Task InvalidTokens()
    {
        await using var host = await TestServerHost.Start();
        await File.WriteAllTextAsync(host.PathOf("/test.docx"), "content");
        var (fileId, token, _) = await host.Open("/test.docx");

        using var invalid = await host.Client.GetAsync($"/wopi/files/{fileId}?access_token={token}x");
        await Assert.That(invalid.StatusCode).IsEqualTo(HttpStatusCode.Unauthorized);
        using var missing = await host.Client.GetAsync($"/wopi/files/{fileId}");
        await Assert.That(missing.StatusCode).IsEqualTo(HttpStatusCode.Unauthorized);

        using var request = new HttpRequestMessage(HttpMethod.Get, $"/wopi/files/{fileId}?access_token={token}");
        request.Headers.Add("X-WOPI-TimeStamp", DateTime.UtcNow.AddMinutes(-30).Ticks.ToString());
        using var expired = await host.Client.SendAsync(request);
        await Assert.That(expired.StatusCode).IsEqualTo(HttpStatusCode.InternalServerError);
        await Assert.That(await Body(expired)).IsEqualTo("Invalid or expired X-WOPI-Timestamp header");
    }

    [Test]
    public async Task GetFile_And_Download()
    {
        await using var host = await TestServerHost.Start();
        await File.WriteAllTextAsync(host.PathOf("/my file.docx"), "content");
        var (fileId, token, _) = await host.Open("/my file.docx");
        using var response = await host.Client.GetAsync($"/wopi/files/{fileId}/contents?access_token={token}");
        await Assert.That(await Body(response)).IsEqualTo("content");
        await Assert.That(response.Content.Headers.ContentDisposition!.ToString()).IsEqualTo("attachment; filename*=UTF-8''my%20file.docx");
        await Assert.That(Header(response, "X-WOPI-ItemVersion")).StartsWith("v");

        using var download = await host.Client.GetAsync($"/wopi/iop/download?access_token={token}");
        await Assert.That(await Body(download)).IsEqualTo("content");
    }

    [Test]
    public async Task Lock_Lifecycle()
    {
        await using var host = await TestServerHost.Start();
        await File.WriteAllTextAsync(host.PathOf("/test.docx"), "content");
        var (fileId, token, _) = await host.Open("/test.docx");

        using (var getLock = await host.Post(fileId, token, "GET_LOCK"))
        {
            await Assert.That(getLock.StatusCode).IsEqualTo(HttpStatusCode.NotFound);
        }

        using (var putUnlocked = await host.PutFile(fileId, token, "new", ("X-WOPI-Lock", "L1")))
        {
            await Assert.That(putUnlocked.StatusCode).IsEqualTo(HttpStatusCode.Conflict);
            await Assert.That(Header(putUnlocked, "X-WOPI-LockFailureReason")).IsEqualTo("Cannot overwrite unlocked file");
        }

        using (var setLock = await host.Post(fileId, token, "LOCK", ("X-WOPI-Lock", "L1")))
        {
            await Assert.That(setLock.StatusCode).IsEqualTo(HttpStatusCode.OK);
            await Assert.That(Header(setLock, "X-WOPI-ItemVersion")).StartsWith("v");
        }

        using (var getLock = await host.Post(fileId, token, "GET_LOCK"))
        {
            await Assert.That(getLock.StatusCode).IsEqualTo(HttpStatusCode.OK);
            await Assert.That(Header(getLock, "X-WOPI-Lock")).IsEqualTo("L1");
        }

        using (var conflict = await host.Post(fileId, token, "LOCK", ("X-WOPI-Lock", "L2"), ("X-WOPI-SessionId", "s2")))
        {
            await Assert.That(conflict.StatusCode).IsEqualTo(HttpStatusCode.Conflict);
            await Assert.That(Header(conflict, "X-WOPI-Lock")).IsEqualTo("L1");
            await Assert.That(await Body(conflict)).IsEqualTo("""{"message":"The file is locked by MSOffice"}""");
        }

        using (var refresh = await host.Post(fileId, token, "REFRESH_LOCK", ("X-WOPI-Lock", "L1")))
        {
            await Assert.That(refresh.StatusCode).IsEqualTo(HttpStatusCode.OK);
        }

        using (var relock = await host.Post(fileId, token, "LOCK", ("X-WOPI-Lock", "L2"), ("X-WOPI-OldLock", "L1")))
        {
            await Assert.That(relock.StatusCode).IsEqualTo(HttpStatusCode.OK);
        }

        using (var putFile = await host.PutFile(fileId, token, "new content", ("X-WOPI-Lock", "L2")))
        {
            await Assert.That(putFile.StatusCode).IsEqualTo(HttpStatusCode.OK);
        }

        await Assert.That(await File.ReadAllTextAsync(host.PathOf("/test.docx"))).IsEqualTo("new content");

        using (var mismatch = await host.PutFile(fileId, token, "lost content", ("X-WOPI-Lock", "L3")))
        {
            await Assert.That(mismatch.StatusCode).IsEqualTo(HttpStatusCode.Conflict);
        }

        // the refused content is kept for recovery
        var recovered = Directory.GetFiles(host.RecoveryPath).Single();
        await Assert.That(recovered).EndsWith("_editedby_root_origat_test.docx");
        await Assert.That(await File.ReadAllTextAsync(recovered)).IsEqualTo("lost content");

        using (var unlockMismatch = await host.Post(fileId, token, "UNLOCK", ("X-WOPI-Lock", "L1")))
        {
            await Assert.That(unlockMismatch.StatusCode).IsEqualTo(HttpStatusCode.Conflict);
        }

        using (var unlock = await host.Post(fileId, token, "UNLOCK", ("X-WOPI-Lock", "L2"), ("X-WOPI-SessionId", "s2")))
        {
            await Assert.That(unlock.StatusCode).IsEqualTo(HttpStatusCode.OK);
        }

        using (var getLock = await host.Post(fileId, token, "GET_LOCK"))
        {
            await Assert.That(getLock.StatusCode).IsEqualTo(HttpStatusCode.NotFound);
        }

        using var conflicts = new HttpRequestMessage(HttpMethod.Get, "/wopi/iop/conflicts");
        conflicts.Headers.Add("Authorization", $"Bearer {TestServerHost.IopSecret}");
        using var conflictsResponse = await host.Client.SendAsync(conflicts);
        var json = JsonNode.Parse(await Body(conflictsResponse))!;
        await Assert.That(json["pending"]!.AsObject().Count).IsEqualTo(0);
        await Assert.That((string?) json["resolved"]!["s2"]!["user"]).IsEqualTo("0:0");
        await Assert.That((int?) json["users"]).IsEqualTo(1);
    }

    [Test]
    public async Task RefreshLock_WithoutLock()
    {
        await using var host = await TestServerHost.Start();
        await File.WriteAllTextAsync(host.PathOf("/test.docx"), "content");
        var (fileId, token, _) = await host.Open("/test.docx");
        using var refresh = await host.Post(fileId, token, "REFRESH_LOCK", ("X-WOPI-Lock", "L1"));
        await Assert.That(refresh.StatusCode).IsEqualTo(HttpStatusCode.Conflict);
        await Assert.That(Header(refresh, "X-WOPI-LockFailureReason")).IsEqualTo("The file was not locked");
    }

    [Test]
    public async Task LibreOfficeLock()
    {
        await using var host = await TestServerHost.Start();
        await File.WriteAllTextAsync(host.PathOf("/test.odt"), "content");
        var (fileId, token, _) = await host.Open("/test.odt", appName: "Collabora");
        using (var setLock = await host.Post(fileId, token, "LOCK", ("X-WOPI-Lock", "L1")))
        {
            await Assert.That(setLock.StatusCode).IsEqualTo(HttpStatusCode.OK);
        }

        // a LibreOffice-compatible lock is created for interoperability
        var lockFile = host.PathOf("/.~lock.test.odt#");
        await Assert.That(await File.ReadAllTextAsync(lockFile)).StartsWith(",root via web collaboration,cernbox,");

        using (var unlock = await host.Post(fileId, token, "UNLOCK", ("X-WOPI-Lock", "L1")))
        {
            await Assert.That(unlock.StatusCode).IsEqualTo(HttpStatusCode.OK);
        }

        await Assert.That(File.Exists(lockFile)).IsFalse();

        // a lock held by a LibreOffice desktop user denies the WOPI lock
        await File.WriteAllTextAsync(lockFile, $",Bob,host,{DateTime.Now:dd.MM.yyyy HH:mm},file:///home/bob/.config/libreoffice/4;");
        using (var conflict = await host.Post(fileId, token, "LOCK", ("X-WOPI-Lock", "L1")))
        {
            await Assert.That(conflict.StatusCode).IsEqualTo(HttpStatusCode.Conflict);
            await Assert.That(Header(conflict, "X-WOPI-LockFailureReason")).IsEqualTo("The file is locked by Bob via LibreOffice for Desktop");
        }
    }

    [Test]
    public async Task PutFile_ExternalModification_CreatesConflictCopy()
    {
        await using var host = await TestServerHost.Start();
        await File.WriteAllTextAsync(host.PathOf("/test.docx"), "content");
        var (fileId, token, _) = await host.Open("/test.docx");
        using (var setLock = await host.Post(fileId, token, "LOCK", ("X-WOPI-Lock", "L1")))
        {
            await Assert.That(setLock.StatusCode).IsEqualTo(HttpStatusCode.OK);
        }

        // someone overwrites the file behind the server's back
        File.SetLastWriteTimeUtc(host.PathOf("/test.docx"), DateTime.UtcNow.AddMinutes(5));
        using var putFile = await host.PutFile(fileId, token, "web content", ("X-WOPI-Lock", "L1"));
        await Assert.That(putFile.StatusCode).IsEqualTo(HttpStatusCode.Conflict);
        await Assert.That(Header(putFile, "X-WOPI-LockFailureReason")).IsEqualTo("The file being edited got moved or overwritten, conflict copy created");
        var copy = Directory.GetFiles(host.StoragePath, "test-webconflict-*.docx").Single();
        await Assert.That(await File.ReadAllTextAsync(copy)).IsEqualTo("web content");
        await Assert.That(await File.ReadAllTextAsync(host.PathOf("/test.docx"))).IsEqualTo("content");
    }

    [Test]
    public async Task PutFile_EditNew()
    {
        await using var host = await TestServerHost.Start();
        await File.WriteAllTextAsync(host.PathOf("/new.docx"), "");
        var (fileId, token, _) = await host.Open("/new.docx");
        using (var putFile = await host.PutFile(fileId, token, "first content"))
        {
            await Assert.That(putFile.StatusCode).IsEqualTo(HttpStatusCode.OK);
        }

        await Assert.That(await File.ReadAllTextAsync(host.PathOf("/new.docx"))).IsEqualTo("first content");
        using var again = await host.PutFile(fileId, token, "second content");
        await Assert.That(again.StatusCode).IsEqualTo(HttpStatusCode.Conflict);
        await Assert.That(await Body(again)).IsEqualTo("""{"message":"File exists"}""");
    }

    [Test]
    public async Task ReadOnlyToken_CannotWrite()
    {
        await using var host = await TestServerHost.Start();
        await File.WriteAllTextAsync(host.PathOf("/test.docx"), "content");
        var (fileId, token, _) = await host.Open("/test.docx", "VIEW_MODE_READ_ONLY");
        using var setLock = await host.Post(fileId, token, "LOCK", ("X-WOPI-Lock", "L1"));
        await Assert.That(setLock.StatusCode).IsEqualTo(HttpStatusCode.Unauthorized);
        using var unknown = await host.Client.PostAsync($"/wopi/files/{fileId}?access_token={token}", null);
        await Assert.That(unknown.StatusCode).IsEqualTo(HttpStatusCode.BadRequest);
    }

    [Test]
    public async Task PutRelative_SuggestedTarget()
    {
        await using var host = await TestServerHost.Start();
        await File.WriteAllTextAsync(host.PathOf("/test.docx"), "content");
        await File.WriteAllTextAsync(host.PathOf("/test.pdf"), "existing");
        var (fileId, token, _) = await host.Open("/test.docx");
        using var response = await host.Post(fileId, token, "PUT_RELATIVE", ("X-WOPI-SuggestedTarget", ".pdf"));
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.OK);
        var json = JsonNode.Parse(await Body(response))!;
        // the existing target is not overwritten
        await Assert.That((string?) json["Name"]).IsEqualTo("test_copy.pdf");
        await Assert.That((string?) json["Url"]).StartsWith("http://localhost/wopi/files/local!");
        await Assert.That((string?) json["HostEditUrl"]).StartsWith("https://office/edit?WOPISrc=");
        await Assert.That(File.Exists(host.PathOf("/test_copy.pdf"))).IsTrue();
        await Assert.That(await File.ReadAllTextAsync(host.PathOf("/test.pdf"))).IsEqualTo("existing");

        // the new access token gives access to the new file
        var url = (string) json["Url"]!;
        using var check = await host.Client.GetAsync(url.Replace("http://localhost", ""));
        await Assert.That((string?) JsonNode.Parse(await Body(check))!["BaseFileName"]).IsEqualTo("test_copy.pdf");
    }

    [Test]
    public async Task PutRelative_RelativeTarget()
    {
        await using var host = await TestServerHost.Start();
        await File.WriteAllTextAsync(host.PathOf("/test.docx"), "content");
        await File.WriteAllTextAsync(host.PathOf("/other.docx"), "existing");
        var (fileId, token, _) = await host.Open("/test.docx");
        using (var conflict = await host.Post(fileId, token, "PUT_RELATIVE", ("X-WOPI-RelativeTarget", "other.docx")))
        {
            await Assert.That(conflict.StatusCode).IsEqualTo(HttpStatusCode.Conflict);
            await Assert.That((string?) JsonNode.Parse(await Body(conflict))!["Name"]).IsEqualTo("other.docx");
        }

        using (var overwrite = await host.Post(fileId, token, "PUT_RELATIVE", ("X-WOPI-RelativeTarget", "other.docx"), ("X-WOPI-OverwriteRelativeTarget", "true")))
        {
            await Assert.That(overwrite.StatusCode).IsEqualTo(HttpStatusCode.OK);
        }

        using var both = await host.Post(fileId, token, "PUT_RELATIVE", ("X-WOPI-RelativeTarget", "a.docx"), ("X-WOPI-SuggestedTarget", "b.docx"));
        await Assert.That(both.StatusCode).IsEqualTo(HttpStatusCode.BadRequest);
    }

    [Test]
    public async Task RenameFile()
    {
        await using var host = await TestServerHost.Start("enablerename = True");
        await File.WriteAllTextAsync(host.PathOf("/test.docx"), "content");
        var (fileId, token, _) = await host.Open("/test.docx");
        using (var info = await host.Client.GetAsync($"/wopi/files/{fileId}?access_token={token}"))
        {
            await Assert.That((bool?) JsonNode.Parse(await Body(info))!["UserCanRename"]).IsTrue();
        }

        using var rename = await host.Post(fileId, token, "RENAME_FILE", ("X-WOPI-RequestedName", "renamed"));
        await Assert.That(rename.StatusCode).IsEqualTo(HttpStatusCode.OK);
        await Assert.That(await Body(rename)).IsEqualTo("""{"Name":"renamed"}""");
        await Assert.That(File.Exists(host.PathOf("/renamed.docx"))).IsTrue();
        await Assert.That(File.Exists(host.PathOf("/test.docx"))).IsFalse();
    }

    [Test]
    public async Task DeleteFile()
    {
        await using var host = await TestServerHost.Start();
        await File.WriteAllTextAsync(host.PathOf("/test.docx"), "content");
        var (fileId, token, _) = await host.Open("/test.docx");
        using (var setLock = await host.Post(fileId, token, "LOCK", ("X-WOPI-Lock", "L1")))
        {
            await Assert.That(setLock.StatusCode).IsEqualTo(HttpStatusCode.OK);
        }

        using (var locked = await host.Post(fileId, token, "DELETE"))
        {
            await Assert.That(locked.StatusCode).IsEqualTo(HttpStatusCode.Conflict);
        }

        using (var unlock = await host.Post(fileId, token, "UNLOCK", ("X-WOPI-Lock", "L1")))
        {
            await Assert.That(unlock.StatusCode).IsEqualTo(HttpStatusCode.OK);
        }

        using var delete = await host.Post(fileId, token, "DELETE");
        await Assert.That(delete.StatusCode).IsEqualTo(HttpStatusCode.OK);
        await Assert.That(File.Exists(host.PathOf("/test.docx"))).IsFalse();
    }

    [Test]
    public async Task PutUserInfo()
    {
        await using var host = await TestServerHost.Start();
        await File.WriteAllTextAsync(host.PathOf("/test.docx"), "content");
        var (fileId, token, _) = await host.Open("/test.docx");
        using (var request = new HttpRequestMessage(HttpMethod.Post, $"/wopi/files/{fileId}?access_token={token}"))
        {
            request.Headers.Add("X-WOPI-Override", "PUT_USER_INFO");
            request.Content = new StringContent("user info");
            using var response = await host.Client.SendAsync(request);
            await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.OK);
        }

        using var info = await host.Client.GetAsync($"/wopi/files/{fileId}?access_token={token}");
        await Assert.That((string?) JsonNode.Parse(await Body(info))!["UserInfo"]).IsEqualTo("user info");
    }

    [Test]
    public async Task OpenFilesList()
    {
        await using var host = await TestServerHost.Start();
        await File.WriteAllTextAsync(host.PathOf("/test.docx"), "content");
        var (fileId, token, _) = await host.Open("/test.docx");
        using (var setLock = await host.Post(fileId, token, "LOCK", ("X-WOPI-Lock", "L1"), ("X-WOPI-SessionId", "s1")))
        {
            await Assert.That(setLock.StatusCode).IsEqualTo(HttpStatusCode.OK);
        }

        using (var unauthorized = await host.Client.GetAsync("/wopi/iop/list"))
        {
            await Assert.That(unauthorized.StatusCode).IsEqualTo(HttpStatusCode.Unauthorized);
        }

        using var request = new HttpRequestMessage(HttpMethod.Get, "/wopi/iop/list");
        request.Headers.Add("Authorization", $"Bearer {TestServerHost.IopSecret}");
        using var response = await host.Client.SendAsync(request);
        var json = JsonNode.Parse(await Body(response))!;
        await Assert.That((string?) json["/test.docx"]![1]![0]).IsEqualTo("s1");
    }

    [Test]
    public async Task WopiValidatorTest()
    {
        await using var host = await TestServerHost.Start();
        await File.WriteAllTextAsync(host.PathOf("/test.wopitest"), "");
        using var request = new HttpRequestMessage(HttpMethod.Get, "/wopi/iop/test?filepath=/test.wopitest&usertoken=0:0");
        request.Headers.Add("Authorization", $"Bearer {TestServerHost.IopSecret}");
        using var response = await host.Client.SendAsync(request);
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.OK);
        await Assert.That(await Body(response)).StartsWith("-e WOPI_URL=http://localhost:8880/wopi/files/local!");
    }
}
