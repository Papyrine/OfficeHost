/// <summary>
/// Edge cases ported from the integration and conformance tests of WopiHost (https://github.com/petrsvihlik/WopiHost).
/// Each test names the WopiHost test it comes from. Several go beyond [MS-WOPI] and follow the Microsoft WOPI
/// documentation (https://learn.microsoft.com/en-us/microsoft-365/cloud-storage-partner-program/rest/), as WopiHost does.
/// </summary>
public class EdgeCaseTests
{
    const string databuf = "file content";

    static string? Header(HttpResponseMessage response, string name)
    {
        if (response.Headers.TryGetValues(name, out var values))
        {
            return string.Join(", ", values);
        }

        return null;
    }

    static async Task<(TestServerHost Host, string FileId, string Token)> Open(string content = databuf, TimeProvider? time = null, string file = "/test.docx", Func<IWopiStorage, IWopiStorage>? wrapStorage = null)
    {
        var host = await TestServerHost.Start(time: time, wrapStorage: wrapStorage);
        await File.WriteAllTextAsync(host.PathOf(file), content);
        var (fileId, token) = await host.Open(file);
        return (host, fileId, token);
    }

    static async Task Lock(TestServerHost host, string fileId, string token, string value)
    {
        using var response = await host.Post(fileId, token, "LOCK", ("X-WOPI-Lock", value));
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.OK);
    }

    static async Task<string?> CurrentLock(TestServerHost host, string fileId, string token)
    {
        using var response = await host.Post(fileId, token, "GET_LOCK");
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.OK);
        return Header(response, "X-WOPI-Lock");
    }

    // Lock identifiers

    /// <summary>
    /// Lock_MissingNewLockIdentifier_Returns_400, RefreshLock_MissingIdentifier_Returns_400,
    /// UnlockAndRelock_MissingNewLock_Returns_400: X-WOPI-Lock is required.
    /// </summary>
    [Test]
    [Arguments("LOCK", null)]
    [Arguments("LOCK", "")]
    [Arguments("REFRESH_LOCK", null)]
    [Arguments("UNLOCK_AND_RELOCK", null)]
    public async Task Lock_MissingLockId(string operation, string? value)
    {
        var (host, fileId, token) = await Open();
        await using var _ = host;
        await Lock(host, fileId, token, "L");
        var headers = new List<(string, string)>();
        if (operation == "UNLOCK_AND_RELOCK")
        {
            operation = "LOCK";
            headers.Add(("X-WOPI-OldLock", "L"));
        }

        if (value != null)
        {
            headers.Add(("X-WOPI-Lock", value));
        }

        using var response = await host.Post(fileId, token, operation, headers.ToArray());
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.BadRequest);
        await Assert.That(await CurrentLock(host, fileId, token)).IsEqualTo("L");
    }

    /// <summary>
    /// Lock_IdTooLong_Returns_400: lock ids are at most 1024 characters (SupportsExtendedLockLength).
    /// </summary>
    [Test]
    [Arguments("LOCK")]
    [Arguments("REFRESH_LOCK")]
    public async Task Lock_IdTooLong(string operation)
    {
        var (host, fileId, token) = await Open();
        await using var _ = host;
        using var response = await host.Post(fileId, token, operation, ("X-WOPI-Lock", new('z', 1025)));
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.BadRequest);
        await Assert.That(string.IsNullOrEmpty(Header(response, "X-WOPI-LockFailureReason"))).IsFalse();
        await Assert.That(await CurrentLock(host, fileId, token)).IsEqualTo("");
    }

    /// <summary>
    /// AddLockAsync_PreservesExoticLockIdBytes_OnGet, TryUnlockAndRelockAsync_PreservesExoticLockIdBytes_OnSwap:
    /// the lock id is opaque, so it round trips whatever its content.
    /// </summary>
    [Test]
    [Arguments("""{"S":"1d0f6c5e","F":4,"E":2,"M":"quoted \"value\""}""")]
    [Arguments("with spaces, commas; and = signs")]
    [Arguments("opaquelocktoken:797356a8-0500-4ceb-a8a0-c94c8cde7eba")]
    public async Task Lock_ExoticIdsRoundTrip(string value)
    {
        var (host, fileId, token) = await Open();
        await using var _ = host;
        await Lock(host, fileId, token, value);
        await Assert.That(await CurrentLock(host, fileId, token)).IsEqualTo(value);
        using (var relock = await host.Post(fileId, token, "LOCK", ("X-WOPI-Lock", $"{value}2"), ("X-WOPI-OldLock", value)))
        {
            await Assert.That(relock.StatusCode).IsEqualTo(HttpStatusCode.OK);
        }

        await Assert.That(await CurrentLock(host, fileId, token)).IsEqualTo($"{value}2");
    }

    /// <summary>
    /// RefreshLockAsync_AfterClockAdvancesPastExpiry_ReturnsFalseAndEvicts,
    /// TryUnlockAndRelockAsync_AfterClockAdvancesPastExpiry_ReturnsFalse: an expired lock cannot be refreshed or swapped.
    /// </summary>
    [Test]
    public async Task Lock_ExpiredCannotBeRefreshedOrSwapped()
    {
        var time = new ManualTime();
        var (host, fileId, token) = await Open(time: time);
        await using var _ = host;
        await Lock(host, fileId, token, "L");
        time.Advance(TimeSpan.FromMinutes(31));
        using (var refresh = await host.Post(fileId, token, "REFRESH_LOCK", ("X-WOPI-Lock", "L")))
        {
            await Assert.That(refresh.StatusCode).IsEqualTo(HttpStatusCode.Conflict);
        }

        using (var relock = await host.Post(fileId, token, "LOCK", ("X-WOPI-Lock", "L2"), ("X-WOPI-OldLock", "L")))
        {
            await Assert.That(relock.StatusCode).IsEqualTo(HttpStatusCode.Conflict);
        }

        await Assert.That(await CurrentLock(host, fileId, token)).IsEqualTo("");
    }

    // PutFile

    /// <summary>
    /// PutFile_OnLockedFile_MissingLockHeader_NonEmpty_Returns_409_WithCurrentLock,
    /// PutFile_OnLockedFile_MissingLockHeader_ZeroByte_Returns_409_WithCurrentLock: a PutFile without lock (the
    /// editnew case) on a locked file is a conflict, which gives the current lock, even on an empty file.
    /// </summary>
    [Test]
    [Arguments(databuf)]
    [Arguments("")]
    public async Task PutFile_LockedFile_MissingLockHeader(string content)
    {
        var (host, fileId, token) = await Open(content);
        await using var _ = host;
        await Lock(host, fileId, token, "L");
        using var response = await host.PutFile(fileId, token, "create-new-attempt");
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.Conflict);
        await Assert.That(Header(response, "X-WOPI-Lock")).IsEqualTo("L");
        await Assert.That(await File.ReadAllTextAsync(host.PathOf("/test.docx"))).IsEqualTo(content);
    }

    /// <summary>
    /// PutFile_StaleBinding_FileRemovedOutOfBand_Returns_404: the file was deleted behind the server's back while locked.
    /// </summary>
    [Test]
    public async Task PutFile_FileRemovedOutOfBand()
    {
        var (host, fileId, token) = await Open();
        await using var _ = host;
        await Lock(host, fileId, token, "L");
        File.Delete(host.PathOf("/test.docx"));
        using var response = await host.PutFile(fileId, token, "update", ("X-WOPI-Lock", "L"));
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.NotFound);
        await Assert.That(File.Exists(host.PathOf("/test.docx"))).IsFalse();
    }

    // GetFile

    /// <summary>
    /// GetFile_Returns_412_WhenSizeExceedsMax: X-WOPI-MaxExpectedSize is an upper bound of the expected size.
    /// </summary>
    [Test]
    [Arguments("1", HttpStatusCode.PreconditionFailed)]
    [Arguments("12", HttpStatusCode.OK)]
    [Arguments("not a number", HttpStatusCode.OK)]
    public async Task GetFile_MaxExpectedSize(string maxExpectedSize, HttpStatusCode expected)
    {
        var (host, fileId, token) = await Open();
        await using var _ = host;
        using var request = new HttpRequestMessage(HttpMethod.Get, $"/wopi/files/{fileId}/contents?access_token={token}");
        request.Headers.Add("X-WOPI-MaxExpectedSize", maxExpectedSize);
        using var response = await host.Client.SendAsync(request);
        await Assert.That(response.StatusCode).IsEqualTo(expected);
    }

    /// <summary>
    /// The storage of a file that vanishes between GetFile and OpenRead.
    /// </summary>
    sealed class VanishingStorage(IWopiStorage inner, Action vanish) :
        ForwardingStorage(inner)
    {
        public override Task<Stream> OpenRead(string fileId, Cancel cancel = default)
        {
            vanish();
            return base.OpenRead(fileId, cancel);
        }
    }

    /// <summary>
    /// GetFile_FileVanishesAfterExistsCheck_Returns_404.
    /// </summary>
    [Test]
    public async Task GetFile_FileVanishesBeforeRead()
    {
        TestServerHost? host = null;
        await using var hostScope = host = await TestServerHost.Start(wrapStorage: _ => new VanishingStorage(_, () => File.Delete(host!.PathOf("/test.docx"))));
        await File.WriteAllTextAsync(host.PathOf("/test.docx"), databuf);
        var (fileId, token) = await host.Open("/test.docx");
        using var response = await host.Client.GetAsync($"/wopi/files/{fileId}/contents?access_token={token}");
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.NotFound);
    }

    // PutUserInfo

    /// <summary>
    /// PutUserInfo_BodyOverLimit_Returns_400: the user info is at most 1024 characters.
    /// </summary>
    [Test]
    [Arguments(1024, HttpStatusCode.OK)]
    [Arguments(1025, HttpStatusCode.BadRequest)]
    public async Task PutUserInfo_Limit(int length, HttpStatusCode expected)
    {
        var (host, fileId, token) = await Open();
        await using var _ = host;
        using var request = new HttpRequestMessage(HttpMethod.Post, $"/wopi/files/{fileId}?access_token={token}");
        request.Headers.Add("X-WOPI-Override", "PUT_USER_INFO");
        request.Content = new StringContent(new('x', length));
        using var response = await host.Client.SendAsync(request);
        await Assert.That(response.StatusCode).IsEqualTo(expected);
    }

    /// <summary>
    /// PutUserInfo_NotFound_Returns_404.
    /// </summary>
    [Test]
    public async Task PutUserInfo_UnknownFile()
    {
        var (host, fileId, token) = await Open();
        await using var _ = host;
        File.Delete(host.PathOf("/test.docx"));
        using var request = new HttpRequestMessage(HttpMethod.Post, $"/wopi/files/{fileId}?access_token={token}");
        request.Headers.Add("X-WOPI-Override", "PUT_USER_INFO");
        request.Content = new StringContent("info");
        using var response = await host.Client.SendAsync(request);
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.NotFound);
    }

    // File names

    static async Task<HttpResponseMessage> PutRelative(TestServerHost host, string fileId, string token, string header, string target)
    {
        using var request = new HttpRequestMessage(HttpMethod.Post, $"/wopi/files/{fileId}?access_token={token}");
        request.Headers.Add("X-WOPI-Override", "PUT_RELATIVE");
        request.Headers.TryAddWithoutValidation(header, target);
        request.Content = new StringContent("body");
        return await host.Client.SendAsync(request);
    }

    /// <summary>
    /// PutRelativeFile_Relative_InvalidName_Returns_400_WithValidRelativeTarget: the 400 for an illegal relative
    /// target suggests a legal name.
    /// </summary>
    [Test]
    [Arguments("bad/name.txt")]
    [Arguments("bad:name.txt")]
    [Arguments("what?.txt")]
    [Arguments("a<b>c.txt")]
    [Arguments("pipe|star*.txt")]
    [Arguments("quote\".txt")]
    [Arguments("trailing.txt.")]
    public async Task PutRelativeFile_RelativeTarget_IllegalName(string target)
    {
        var (host, fileId, token) = await Open();
        await using var _ = host;
        using var response = await PutRelative(host, fileId, token, "X-WOPI-RelativeTarget", target);
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.BadRequest);
        var valid = FileNames.DecodeUtf7(Header(response, "X-WOPI-ValidRelativeTarget") ?? "");
        await Assert.That(valid).IsNotEqualTo("");
        await Assert.That(valid.IndexOfAny(['/', '\\', ':', '?', '<', '>', '|', '*', '"'])).IsEqualTo(-1);
        await Assert.That(valid).EndsWith(".txt");

        using var retry = await PutRelative(host, fileId, token, "X-WOPI-RelativeTarget", FileNames.EncodeUtf7(valid));
        await Assert.That(retry.StatusCode).IsEqualTo(HttpStatusCode.OK);
    }

    /// <summary>
    /// PutRelativeFile_Suggested_InvalidName_Returns_200_NotBadRequest, ScrubOrNull_ReplacesForbiddenChars: a
    /// suggested name is scrubbed of the characters file systems forbid, keeping its extension.
    /// </summary>
    [Test]
    [Arguments("bad:name?.txt")]
    [Arguments("a<b>c|d*e\".txt")]
    [Arguments("trailing .txt")]
    public async Task PutRelativeFile_SuggestedTarget_Scrubbed(string target)
    {
        var (host, fileId, token) = await Open();
        await using var _ = host;
        using var response = await PutRelative(host, fileId, token, "X-WOPI-SuggestedTarget", target);
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.OK);
        var name = (string) JsonNode.Parse(await response.Content.ReadAsStringAsync())!["Name"]!;
        await Assert.That(name.IndexOfAny([':', '?', '<', '>', '|', '*', '"'])).IsEqualTo(-1);
        await Assert.That(name).EndsWith(".txt");
        await Assert.That(await File.ReadAllTextAsync(host.PathOf($"/{name}"))).IsEqualTo("body");
    }

    /// <summary>
    /// ScrubOrNull_ReturnsNull_ForEmptyOrPathNav, TryBuildValidCandidateAsync_EmptyOrPathNavStem_SubstitutesFallback:
    /// a suggested name with nothing usable left still gives a legal file.
    /// </summary>
    [Test]
    [Arguments("..")]
    [Arguments("/")]
    [Arguments("?.txt")]
    public async Task PutRelativeFile_SuggestedTarget_NothingLeft(string target)
    {
        var (host, fileId, token) = await Open();
        await using var _ = host;
        using var response = await PutRelative(host, fileId, token, "X-WOPI-SuggestedTarget", target);
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.OK);
        var name = (string) JsonNode.Parse(await response.Content.ReadAsStringAsync())!["Name"]!;
        await Assert.That(name is "" or "." or "..").IsFalse();
        await Assert.That(File.Exists(host.PathOf($"/{name}"))).IsTrue();
    }

    static string longName = string.Concat(Enumerable.Repeat("longwopitestfilename", 26));

    /// <summary>
    /// WOPI validator PutRelativeFile.FileNameLongerThan512Chars: a host supports long names or returns a 400. Names
    /// are capped to 250 UTF-8 bytes, the default FileNameMaxLength of [MS-WOPI].
    /// </summary>
    [Test]
    [Arguments("X-WOPI-RelativeTarget", ".wopitest", HttpStatusCode.BadRequest)]
    [Arguments("X-WOPI-SuggestedTarget", ".wopitest", HttpStatusCode.OK)]
    [Arguments("X-WOPI-SuggestedTarget", ".txt", HttpStatusCode.OK)]
    public async Task PutRelativeFile_LongName(string header, string extension, HttpStatusCode expected)
    {
        var (host, fileId, token) = await Open();
        await using var _ = host;
        using var response = await PutRelative(host, fileId, token, header, $"{longName}{extension}");
        await Assert.That(response.StatusCode).IsEqualTo(expected);
        string name;
        if (expected == HttpStatusCode.OK)
        {
            name = (string) JsonNode.Parse(await response.Content.ReadAsStringAsync())!["Name"]!;
            await Assert.That(File.Exists(host.PathOf($"/{name}"))).IsTrue();
        }
        else
        {
            name = FileNames.DecodeUtf7(Header(response, "X-WOPI-ValidRelativeTarget")!);
        }

        await Assert.That(Encoding.UTF8.GetByteCount(name)).IsLessThanOrEqualTo(250);
        await Assert.That(name).EndsWith(extension);
        await Assert.That(name).StartsWith("longwopitestfilename");
    }

    /// <summary>
    /// Names are truncated by UTF-8 bytes, without splitting a character.
    /// </summary>
    [Test]
    public async Task PutRelativeFile_LongMultiByteName()
    {
        var (host, fileId, token) = await Open();
        await using var _ = host;
        using var response = await PutRelative(host, fileId, token, "X-WOPI-SuggestedTarget", FileNames.EncodeUtf7($"{new string('é', 200)}😀.docx"));
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.OK);
        var name = (string) JsonNode.Parse(await response.Content.ReadAsStringAsync())!["Name"]!;
        await Assert.That(Encoding.UTF8.GetByteCount(name)).IsLessThanOrEqualTo(250);
        await Assert.That(name).EndsWith(".docx");
        await Assert.That(name.Contains('�')).IsFalse();
    }

    [Test]
    public async Task RenameFile_LongName()
    {
        var (host, fileId, token) = await Open();
        await using var _ = host;
        using var response = await host.Post(fileId, token, "RENAME_FILE", ("X-WOPI-Lock", ""), ("X-WOPI-RequestedName", longName));
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.BadRequest);
    }

    /// <summary>
    /// RenameFile_InvalidName: names with characters file systems forbid are illegal. WopiHost scrubs them and returns
    /// 200; this server follows the WOPI documentation, a 400 with X-WOPI-InvalidFileNameError.
    /// </summary>
    [Test]
    [Arguments("bad:name")]
    [Arguments("what?")]
    [Arguments("trailing.")]
    public async Task RenameFile_ForbiddenCharacters(string name)
    {
        var (host, fileId, token) = await Open();
        await using var _ = host;
        using var response = await host.Post(fileId, token, "RENAME_FILE", ("X-WOPI-Lock", ""), ("X-WOPI-RequestedName", name));
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.BadRequest);
        await Assert.That(string.IsNullOrEmpty(Header(response, "X-WOPI-InvalidFileNameError"))).IsFalse();
        await Assert.That(File.Exists(host.PathOf("/test.docx"))).IsTrue();
    }
}
