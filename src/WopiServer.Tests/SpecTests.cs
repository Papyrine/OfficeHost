/// <summary>
/// Conformance to the server side of [MS-WOPI] (Web Application Open Platform Interface Protocol, version of 2025-08-19).
/// Each test cites the section it checks. The operations the server does not implement are checked to be reported as
/// unsupported, see <see cref="Unsupported_Operations"/>.
/// </summary>
public class SpecTests
{
    const string databuf = "file content";

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

    static async Task<(TestServerHost Host, string FileId, string Token)> Open(string file = "/test.docx", string content = databuf, Mode mode = Mode.Edit, Action<WopiOptions>? configure = null, TimeProvider? time = null)
    {
        var host = await TestServerHost.Start(configure, time);
        await File.WriteAllTextAsync(host.PathOf(file), content);
        var (fileId, token) = await host.Open(file, mode);
        return (host, fileId, token);
    }

    static async Task<JsonObject> CheckFileInfo(TestServerHost host, string fileId, string token)
    {
        using var response = await host.Client.GetAsync($"/wopi/files/{fileId}?access_token={token}");
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.OK);
        return JsonNode.Parse(await Body(response))!.AsObject();
    }

    static async Task<string?> CurrentLock(TestServerHost host, string fileId, string token)
    {
        using var response = await host.Post(fileId, token, "GET_LOCK");
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.OK);
        return Header(response, "X-WOPI-Lock");
    }

    static async Task Lock(TestServerHost host, string fileId, string token, string value)
    {
        using var response = await host.Post(fileId, token, "LOCK", ("X-WOPI-Lock", value));
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.OK);
    }

    static async Task<string> Version(TestServerHost host, string fileId, string token) =>
        (string) (await CheckFileInfo(host, fileId, token))["Version"]!;

    // 2.2.1 Custom HTTP Headers

    /// <summary>
    /// 2.2.1: X-WOPI-ServerVersion and X-WOPI-MachineName MUST be included with all WOPI responses.
    /// </summary>
    [Test]
    public async Task Responses_IncludeServerVersionAndMachineName()
    {
        var (host, fileId, token) = await Open();
        await using var _ = host;
        using var checkFileInfo = await host.Client.GetAsync($"/wopi/files/{fileId}?access_token={token}");
        using var getFile = await host.Client.GetAsync($"/wopi/files/{fileId}/contents?access_token={token}");
        using var invalidToken = await host.Client.GetAsync($"/wopi/files/{fileId}?access_token=invalid");
        using var conflict = await host.PutFile(fileId, token, "x", ("X-WOPI-Lock", "unknown"));
        foreach (var response in new[] { checkFileInfo, getFile, invalidToken, conflict })
        {
            await Assert.That(string.IsNullOrEmpty(Header(response, "X-WOPI-ServerVersion"))).IsFalse();
            await Assert.That(string.IsNullOrEmpty(Header(response, "X-WOPI-MachineName"))).IsFalse();
        }
    }

    /// <summary>
    /// 2.2.1: X-WOPI-ServerError is included in a WOPI response if the status code is 500.
    /// </summary>
    [Test]
    public async Task ServerError_IncludesServerErrorHeader()
    {
        var (host, fileId, token) = await Open();
        await using var _ = host;
        using var request = new HttpRequestMessage(HttpMethod.Get, $"/wopi/files/{fileId}?access_token={token}");
        // an expired X-WOPI-TimeStamp is answered with a 500, as the WOPI validator requires
        request.Headers.Add("X-WOPI-TimeStamp", DateTime.UtcNow.AddHours(-1).Ticks.ToString());
        using var response = await host.Client.SendAsync(request);
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.InternalServerError);
        await Assert.That(string.IsNullOrEmpty(Header(response, "X-WOPI-ServerError"))).IsFalse();
    }

    /// <summary>
    /// 2.2.1: all WOPI requests include the token as "Bearer " + token in the Authorization header,
    /// besides the access_token URI parameter of 3.3.5.
    /// </summary>
    [Test]
    public async Task Requests_WithAuthorizationHeader()
    {
        var (host, fileId, token) = await Open();
        await using var _ = host;
        using var request = new HttpRequestMessage(HttpMethod.Get, $"/wopi/files/{fileId}?access_token={token}");
        request.Headers.Add("Authorization", $"Bearer {token}");
        using var response = await host.Client.SendAsync(request);
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.OK);
    }

    // 2.2.3 Common URI Parameters

    /// <summary>
    /// 2.2.3: the server MUST issue a unique id for any resource, and 3.3.5.1 requires it to be URL safe.
    /// The id must also not change when the file is saved, as WOPI clients keep using it.
    /// </summary>
    [Test]
    public async Task FileIds_AreUniqueUrlSafeAndStable()
    {
        await using var host = await TestServerHost.Start();
        await File.WriteAllTextAsync(host.PathOf("/a.docx"), databuf);
        await File.WriteAllTextAsync(host.PathOf("/b.docx"), databuf);
        var (idA, tokenA) = await host.Open("/a.docx");
        var (idB, _) = await host.Open("/b.docx");
        await Assert.That(idA).IsNotEqualTo(idB);
        foreach (var id in new[] { idA, idB })
        {
            // RFC 3986 unreserved and sub-delims characters need no escaping in a path segment
            await Assert.That(id.All(_ => char.IsAsciiLetterOrDigit(_) || "-._~!$&'()*+,;=".Contains(_))).IsTrue();
        }

        await Lock(host, idA, tokenA, "L");
        using (var putFile = await host.PutFile(idA, tokenA, "new content", ("X-WOPI-Lock", "L")))
        {
            await Assert.That(putFile.StatusCode).IsEqualTo(HttpStatusCode.OK);
        }

        var (idAfterSave, _) = await host.Open("/a.docx");
        await Assert.That(idAfterSave).IsEqualTo(idA);
    }

    /// <summary>
    /// 2.2.3: the token SHOULD be scoped to a specific user and set of resources.
    /// </summary>
    [Test]
    public async Task Tokens_AreScopedToTheirFile()
    {
        await using var host = await TestServerHost.Start();
        await File.WriteAllTextAsync(host.PathOf("/a.docx"), databuf);
        await File.WriteAllTextAsync(host.PathOf("/b.docx"), "other content");
        var (_, tokenA) = await host.Open("/a.docx");
        var (idB, _) = await host.Open("/b.docx");
        using var response = await host.Client.GetAsync($"/wopi/files/{idB}/contents?access_token={tokenA}");
        await Assert.That(await Body(response)).IsNotEqualTo("other content");
    }

    // 3.3.5.1.2 CheckFileInfo

    /// <summary>
    /// 3.3.5.1.2.2: BaseFileName, OwnerId, Size and Version are required; BaseFileName is the name without
    /// the path, Size is the size in bytes.
    /// </summary>
    [Test]
    public async Task CheckFileInfo_RequiredProperties()
    {
        var (host, fileId, token) = await Open(content: "12345678");
        await using var _ = host;
        var info = await CheckFileInfo(host, fileId, token);
        await Assert.That((string?) info["BaseFileName"]).IsEqualTo("test.docx");
        await Assert.That(string.IsNullOrEmpty((string?) info["OwnerId"])).IsFalse();
        await Assert.That((long?) info["Size"]).IsEqualTo(8);
        await Assert.That(string.IsNullOrEmpty((string?) info["Version"])).IsFalse();
    }

    /// <summary>
    /// 3.3.5.1.2.2 lists OfficeCollaborationServiceEndpointUrl, OpenInClientCommandUrl, OpenInClientPostMessage and
    /// UserPrincipalName as not optional too. They are CSPP+ properties, which break regular Office for the web
    /// sessions: an application that is a CSPP+ partner adds them with OnCheckFileInfo.
    /// </summary>
    [Test]
    public async Task CheckFileInfo_RequiredCsppProperties()
    {
        var (host, fileId, token) = await Open(configure: _ => _.OnCheckFileInfo = AddCsppProperties);
        await using var _ = host;
        var info = await CheckFileInfo(host, fileId, token);
        foreach (var name in CheckFileInfoSchema.Required)
        {
            await Assert.That(info.ContainsKey(name)).IsTrue();
        }
    }

    static Task AddCsppProperties(CheckFileInfoContext context)
    {
        var properties = context.Properties;
        properties["OfficeCollaborationServiceEndpointUrl"] = "";
        properties["OpenInClientCommandUrl"] = "";
        properties["OpenInClientPostMessage"] = false;
        properties["UserPrincipalName"] = $"{context.Access.UserId}@example.com";
        return Task.CompletedTask;
    }

    public static IEnumerable<Func<(string File, Mode Mode, bool Anonymous, bool Customized)>> CheckFileInfoCases()
    {
        yield return () => ("/test.docx", Mode.Edit, false, false);
        yield return () => ("/test.docx", Mode.ViewOnly, true, false);
        yield return () => ("/test.docx", Mode.ReadOnly, false, false);
        yield return () => ("/noextension", Mode.Edit, false, false);
        yield return () => ("/test.docx", Mode.Edit, false, true);
    }

    /// <summary>
    /// 3.3.5.1.2.2: each property has the type given by the schema, whatever the user, the access and the
    /// properties the application adds.
    /// </summary>
    [Test]
    [MethodDataSource(nameof(CheckFileInfoCases))]
    public async Task CheckFileInfo_PropertyTypes(string file, Mode mode, bool anonymous, bool customized)
    {
        await using var host = await TestServerHost.Start(
            _ =>
            {
                if (customized)
                {
                    _.OnCheckFileInfo = AddCsppProperties;
                }
            });
        await File.WriteAllTextAsync(host.PathOf(file), databuf);
        var (fileId, token) = await host.Open(file, mode, anonymous: anonymous);
        var info = await CheckFileInfo(host, fileId, token);
        foreach (var (name, value) in info)
        {
            if (!CheckFileInfoSchema.Types.TryGetValue(name, out var type))
            {
                // extensions for other WOPI clients (Collabora, ONLYOFFICE) and CSPP+ properties are not in the spec
                continue;
            }

            var kind = value?.GetValueKind();
            var matches = type switch
            {
                "string" => kind == JsonValueKind.String,
                "bool" => kind is JsonValueKind.True or JsonValueKind.False,
                "integer" or "long" => kind == JsonValueKind.Number && value!.AsValue().TryGetValue<long>(out _),
                _ => kind == JsonValueKind.Array && value!.AsArray().All(_ => _?.GetValueKind() == JsonValueKind.String)
            };
            if (!matches)
            {
                throw new($"{name} is {kind?.ToString() ?? "null"} but must be {type}");
            }
        }
    }

    /// <summary>
    /// 3.3.5.1.2.2: FileExtension MUST begin with a ".".
    /// </summary>
    [Test]
    [Arguments("/test.docx", ".docx")]
    [Arguments("/archive.tar.gz", ".gz")]
    [Arguments("/noextension", null)]
    public async Task CheckFileInfo_FileExtension(string file, string? expected)
    {
        var (host, fileId, token) = await Open(file);
        await using var _ = host;
        var info = await CheckFileInfo(host, fileId, token);
        await Assert.That((string?) info["FileExtension"]).IsEqualTo(expected);
    }

    /// <summary>
    /// 3.3.5.1.2.2: Version MUST change when the file content changes, and version values MUST never repeat.
    /// </summary>
    [Test]
    public async Task CheckFileInfo_VersionChangesWithContent()
    {
        var (host, fileId, token) = await Open();
        await using var _ = host;
        await Lock(host, fileId, token, "L");
        var versions = new List<string>
        {
            await Version(host, fileId, token)
        };
        foreach (var content in new[] { "one", "two", "one" })
        {
            using var putFile = await host.PutFile(fileId, token, content, ("X-WOPI-Lock", "L"));
            await Assert.That(putFile.StatusCode).IsEqualTo(HttpStatusCode.OK);
            versions.Add(await Version(host, fileId, token));
        }

        await Assert.That(versions.Distinct().Count()).IsEqualTo(versions.Count);
    }

    /// <summary>
    /// 3.3.5.1.2: 404 for an unknown file, 401 for an invalid token.
    /// </summary>
    [Test]
    public async Task CheckFileInfo_Errors()
    {
        var (host, fileId, token) = await Open();
        await using var _ = host;
        using (var invalid = await host.Client.GetAsync($"/wopi/files/{fileId}?access_token={token[..^2]}"))
        {
            await Assert.That(invalid.StatusCode).IsEqualTo(HttpStatusCode.Unauthorized);
        }

        File.Delete(host.PathOf("/test.docx"));
        using var unknown = await host.Client.GetAsync($"/wopi/files/{fileId}?access_token={token}");
        await Assert.That(unknown.StatusCode).IsEqualTo(HttpStatusCode.NotFound);
    }

    /// <summary>
    /// 3.3.5.1.2: X-WOPI-SessionContext can be included with the request.
    /// </summary>
    [Test]
    public async Task CheckFileInfo_SessionContext()
    {
        var (host, fileId, token) = await Open();
        await using var _ = host;
        using var request = new HttpRequestMessage(HttpMethod.Get, $"/wopi/files/{fileId}?access_token={token}");
        request.Headers.Add("X-WOPI-SessionContext", "context");
        using var response = await host.Client.SendAsync(request);
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.OK);
    }

    /// <summary>
    /// 3.3.5.1.2.2: UserCanWrite, ReadOnly and RestrictedWebViewOnly reflect the permissions of the token.
    /// </summary>
    [Test]
    [Arguments(Mode.Edit, true, false)]
    [Arguments(Mode.ReadOnly, false, false)]
    [Arguments(Mode.ViewOnly, false, true)]
    public async Task CheckFileInfo_Permissions(Mode mode, bool canWrite, bool viewOnly)
    {
        var (host, fileId, token) = await Open(mode: mode);
        await using var _ = host;
        var info = await CheckFileInfo(host, fileId, token);
        await Assert.That((bool?) info["UserCanWrite"]).IsEqualTo(canWrite);
        await Assert.That((bool?) info["ReadOnly"]).IsEqualTo(!canWrite);
        await Assert.That((bool?) info["RestrictedWebViewOnly"]).IsEqualTo(viewOnly);
    }

    /// <summary>
    /// 3.3.5.1.2.2: each Supports* property the server sets is backed by the operations it names.
    /// </summary>
    [Test]
    public async Task CheckFileInfo_SupportedOperationsWork()
    {
        var (host, fileId, token) = await Open();
        await using var _ = host;
        var info = await CheckFileInfo(host, fileId, token);
        await Assert.That((bool?) info["SupportsLocks"]).IsTrue();
        await Assert.That((bool?) info["SupportsGetLock"]).IsTrue();
        await Assert.That((bool?) info["SupportsUpdate"]).IsTrue();
        await Assert.That((bool?) info["SupportsUserInfo"]).IsTrue();
        await Assert.That((bool?) info["SupportsDeleteFile"]).IsTrue();

        await Lock(host, fileId, token, "L");
        await Assert.That(await CurrentLock(host, fileId, token)).IsEqualTo("L");
        using (var refresh = await host.Post(fileId, token, "REFRESH_LOCK", ("X-WOPI-Lock", "L")))
        {
            await Assert.That(refresh.StatusCode).IsEqualTo(HttpStatusCode.OK);
        }

        using (var relock = await host.Post(fileId, token, "LOCK", ("X-WOPI-Lock", "L2"), ("X-WOPI-OldLock", "L")))
        {
            await Assert.That(relock.StatusCode).IsEqualTo(HttpStatusCode.OK);
        }

        using (var putFile = await host.PutFile(fileId, token, "x", ("X-WOPI-Lock", "L2")))
        {
            await Assert.That(putFile.StatusCode).IsEqualTo(HttpStatusCode.OK);
        }

        using (var putRelative = await host.Post(fileId, token, "PUT_RELATIVE", ("X-WOPI-SuggestedTarget", ".pdf")))
        {
            await Assert.That(putRelative.StatusCode).IsEqualTo(HttpStatusCode.OK);
        }

        using (var unlock = await host.Post(fileId, token, "UNLOCK", ("X-WOPI-Lock", "L2")))
        {
            await Assert.That(unlock.StatusCode).IsEqualTo(HttpStatusCode.OK);
        }

        using var delete = await host.Post(fileId, token, "DELETE");
        await Assert.That(delete.StatusCode).IsEqualTo(HttpStatusCode.OK);
    }

    /// <summary>
    /// 3.3.5.1.2.2: SupportsExtendedLockLength means lock IDs up to 1024 ASCII characters are supported.
    /// </summary>
    [Test]
    public async Task CheckFileInfo_ExtendedLockLength()
    {
        var (host, fileId, token) = await Open();
        await using var _ = host;
        var info = await CheckFileInfo(host, fileId, token);
        await Assert.That((bool?) info["SupportsExtendedLockLength"]).IsTrue();
        var value = string.Concat(Enumerable.Range(0, 1024).Select(_ => (char) ('!' + _ % 94)));
        await Lock(host, fileId, token, value);
        await Assert.That(await CurrentLock(host, fileId, token)).IsEqualTo(value);
    }

    // 3.3.5.1.16 PutUserInfo

    /// <summary>
    /// 3.3.5.1.16: the server MUST store the string and associate it with the current user, and MUST include it in the
    /// UserInfo property of CheckFileInfo for this user.
    /// </summary>
    [Test]
    public async Task PutUserInfo_IsReturnedForTheSameUserOnly()
    {
        await using var host = await TestServerHost.Start();
        await File.WriteAllTextAsync(host.PathOf("/test.docx"), databuf);
        var (fileId, alice) = await host.Open("/test.docx", userId: "alice");
        var (_, bob) = await host.Open("/test.docx", userId: "bob");
        using (var request = new HttpRequestMessage(HttpMethod.Post, $"/wopi/files/{fileId}?access_token={alice}"))
        {
            request.Headers.Add("X-WOPI-Override", "PUT_USER_INFO");
            request.Content = new StringContent("{\"licensed\":true}");
            using var response = await host.Client.SendAsync(request);
            await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.OK);
        }

        await Assert.That((string?) (await CheckFileInfo(host, fileId, alice))["UserInfo"]).IsEqualTo("{\"licensed\":true}");
        await Assert.That((await CheckFileInfo(host, fileId, bob)).ContainsKey("UserInfo")).IsFalse();
    }

    // 3.3.5.3.1 GetFile

    /// <summary>
    /// 3.3.5.3.1: the server MUST return the complete binary of the file.
    /// </summary>
    [Test]
    public async Task GetFile_ReturnsCompleteBinary()
    {
        await using var host = await TestServerHost.Start();
        var content = new byte[1024 * 1024 + 7];
        new Random(42).NextBytes(content);
        await File.WriteAllBytesAsync(host.PathOf("/test.docx"), content);
        var (fileId, token) = await host.Open("/test.docx");
        using var response = await host.Client.GetAsync($"/wopi/files/{fileId}/contents?access_token={token}");
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.OK);
        await Assert.That((await response.Content.ReadAsByteArrayAsync()).SequenceEqual(content)).IsTrue();
    }

    [Test]
    public async Task GetFile_EmptyFile()
    {
        var (host, fileId, token) = await Open(content: "");
        await using var _ = host;
        using var response = await host.Client.GetAsync($"/wopi/files/{fileId}/contents?access_token={token}");
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.OK);
        await Assert.That(await Body(response)).IsEqualTo("");
    }

    /// <summary>
    /// 3.3.5.3.1: 404 for an unknown file, 401 for an invalid token.
    /// </summary>
    [Test]
    public async Task GetFile_Errors()
    {
        var (host, fileId, token) = await Open();
        await using var _ = host;
        using (var invalid = await host.Client.GetAsync($"/wopi/files/{fileId}/contents?access_token=invalid"))
        {
            await Assert.That(invalid.StatusCode).IsEqualTo(HttpStatusCode.Unauthorized);
        }

        File.Delete(host.PathOf("/test.docx"));
        using var unknown = await host.Client.GetAsync($"/wopi/files/{fileId}/contents?access_token={token}");
        await Assert.That(unknown.StatusCode).IsEqualTo(HttpStatusCode.NotFound);
    }

    // 3.3.5.3.2 PutFile

    /// <summary>
    /// 3.3.5.3.2: with the matching lock, the server MUST update the binary of the file. X-WOPI-ItemVersion SHOULD be
    /// the Version of CheckFileInfo, and X-WOPI-Lock SHOULD NOT be included with a 200.
    /// </summary>
    [Test]
    public async Task PutFile_WithMatchingLock()
    {
        var (host, fileId, token) = await Open();
        await using var _ = host;
        await Lock(host, fileId, token, "L");
        var content = new byte[100_000];
        new Random(1).NextBytes(content);
        using var request = new HttpRequestMessage(HttpMethod.Post, $"/wopi/files/{fileId}/contents?access_token={token}");
        request.Headers.Add("X-WOPI-Override", "PUT");
        request.Headers.Add("X-WOPI-Lock", "L");
        request.Content = new ByteArrayContent(content);
        using var response = await host.Client.SendAsync(request);
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.OK);
        await Assert.That(Header(response, "X-WOPI-Lock")).IsNull();
        await Assert.That((await File.ReadAllBytesAsync(host.PathOf("/test.docx"))).SequenceEqual(content)).IsTrue();
        await Assert.That(Header(response, "X-WOPI-ItemVersion")).IsEqualTo(await Version(host, fileId, token));
    }

    /// <summary>
    /// 3.3.5.3.2: a lock mismatch is a 409, which SHOULD include the current lock.
    /// </summary>
    [Test]
    public async Task PutFile_LockMismatch()
    {
        var (host, fileId, token) = await Open();
        await using var _ = host;
        await Lock(host, fileId, token, "L");
        using var response = await host.PutFile(fileId, token, "x", ("X-WOPI-Lock", "other"));
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.Conflict);
        await Assert.That(Header(response, "X-WOPI-Lock")).IsEqualTo("L");
        await Assert.That(await File.ReadAllTextAsync(host.PathOf("/test.docx"))).IsEqualTo(databuf);
    }

    /// <summary>
    /// 3.3.5.3.2: a file that is not locked, and is not empty, cannot be overwritten by a client that thinks it holds a lock.
    /// </summary>
    [Test]
    public async Task PutFile_UnlockedFile()
    {
        var (host, fileId, token) = await Open();
        await using var _ = host;
        using var response = await host.PutFile(fileId, token, "x", ("X-WOPI-Lock", "L"));
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.Conflict);
        await Assert.That(Header(response, "X-WOPI-Lock")).IsEqualTo("");
        await Assert.That(await File.ReadAllTextAsync(host.PathOf("/test.docx"))).IsEqualTo(databuf);
    }

    /// <summary>
    /// 2.2.3: the token determines the permissions, so a read-only token cannot update the file.
    /// </summary>
    [Test]
    public async Task PutFile_ReadOnlyToken()
    {
        var (host, fileId, token) = await Open(content: "", mode: Mode.ReadOnly);
        await using var _ = host;
        using var response = await host.PutFile(fileId, token, "x");
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.Unauthorized);
        await Assert.That(await File.ReadAllTextAsync(host.PathOf("/test.docx"))).IsEqualTo("");
    }

    [Test]
    public async Task PutFile_InvalidToken()
    {
        var (host, fileId, _) = await Open();
        await using var _ = host;
        using var response = await host.PutFile(fileId, "invalid", "x", ("X-WOPI-Lock", "L"));
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.Unauthorized);
    }

    // 3.3.5.1.14 Lock

    /// <summary>
    /// 3.3.5.1.14: the server MUST create the lock. X-WOPI-Lock SHOULD NOT be included with a 200, and
    /// X-WOPI-ItemVersion should be the Version of CheckFileInfo.
    /// </summary>
    [Test]
    public async Task Lock_UnlockedFile()
    {
        var (host, fileId, token) = await Open();
        await using var _ = host;
        using var response = await host.Post(fileId, token, "LOCK", ("X-WOPI-Lock", "L"));
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.OK);
        await Assert.That(Header(response, "X-WOPI-Lock")).IsNull();
        await Assert.That(Header(response, "X-WOPI-ItemVersion")).IsEqualTo(await Version(host, fileId, token));
        await Assert.That(await CurrentLock(host, fileId, token)).IsEqualTo("L");
    }

    /// <summary>
    /// 3.3.5.1.14: taking the lock already held by the client succeeds.
    /// </summary>
    [Test]
    public async Task Lock_SameLock()
    {
        var (host, fileId, token) = await Open();
        await using var _ = host;
        await Lock(host, fileId, token, "L");
        await Lock(host, fileId, token, "L");
        await Assert.That(await CurrentLock(host, fileId, token)).IsEqualTo("L");
    }

    /// <summary>
    /// 3.3.5.1.14: a lock mismatch is a 409. X-WOPI-Lock SHOULD give the existing lock, and X-WOPI-LockFailureReason
    /// SHOULD be included.
    /// </summary>
    [Test]
    public async Task Lock_DifferentLock()
    {
        var (host, fileId, token) = await Open();
        await using var _ = host;
        await Lock(host, fileId, token, "L");
        using var response = await host.Post(fileId, token, "LOCK", ("X-WOPI-Lock", "other"));
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.Conflict);
        await Assert.That(Header(response, "X-WOPI-Lock")).IsEqualTo("L");
        await Assert.That(string.IsNullOrEmpty(Header(response, "X-WOPI-LockFailureReason"))).IsFalse();
        await Assert.That(await CurrentLock(host, fileId, token)).IsEqualTo("L");
    }

    /// <summary>
    /// 3.3.5.1.14: the lock MUST last for 30 minutes.
    /// </summary>
    [Test]
    public async Task Lock_LastsThirtyMinutes()
    {
        var time = new ManualTime();
        var (host, fileId, token) = await Open(time: time);
        await using var _ = host;
        await Lock(host, fileId, token, "L");
        time.Advance(TimeSpan.FromMinutes(29));
        await Assert.That(await CurrentLock(host, fileId, token)).IsEqualTo("L");
        time.Advance(TimeSpan.FromMinutes(2));
        await Assert.That(await CurrentLock(host, fileId, token)).IsEqualTo("");
    }

    /// <summary>
    /// 3.3.5.1.14: 404 for an unknown file, 401 for an invalid token.
    /// </summary>
    [Test]
    public async Task Lock_Errors()
    {
        var (host, fileId, token) = await Open();
        await using var _ = host;
        using (var invalid = await host.Post(fileId, "invalid", "LOCK", ("X-WOPI-Lock", "L")))
        {
            await Assert.That(invalid.StatusCode).IsEqualTo(HttpStatusCode.Unauthorized);
        }

        File.Delete(host.PathOf("/test.docx"));
        using var unknown = await host.Post(fileId, token, "LOCK", ("X-WOPI-Lock", "L"));
        await Assert.That(unknown.StatusCode).IsEqualTo(HttpStatusCode.NotFound);
    }

    // 3.3.5.1.10 GetLock

    /// <summary>
    /// 3.3.5.1.10: X-WOPI-Lock MUST be included with a 200, and is empty when the file is not locked.
    /// </summary>
    [Test]
    public async Task GetLock_LockedAndUnlocked()
    {
        var (host, fileId, token) = await Open();
        await using var _ = host;
        await Assert.That(await CurrentLock(host, fileId, token)).IsEqualTo("");
        await Lock(host, fileId, token, "L");
        await Assert.That(await CurrentLock(host, fileId, token)).IsEqualTo("L");
    }

    /// <summary>
    /// 3.3.5.1.10: 404 for an unknown file, 401 for an invalid token.
    /// </summary>
    [Test]
    public async Task GetLock_Errors()
    {
        var (host, fileId, token) = await Open();
        await using var _ = host;
        using (var invalid = await host.Post(fileId, "invalid", "GET_LOCK"))
        {
            await Assert.That(invalid.StatusCode).IsEqualTo(HttpStatusCode.Unauthorized);
        }

        File.Delete(host.PathOf("/test.docx"));
        using var unknown = await host.Post(fileId, token, "GET_LOCK");
        await Assert.That(unknown.StatusCode).IsEqualTo(HttpStatusCode.NotFound);
    }

    // 3.3.5.1.18 RefreshLock

    /// <summary>
    /// 3.3.5.1.18: the lock MUST expire 30 minutes after a refresh.
    /// </summary>
    [Test]
    public async Task RefreshLock_RenewsForThirtyMinutes()
    {
        var time = new ManualTime();
        var (host, fileId, token) = await Open(time: time);
        await using var _ = host;
        await Lock(host, fileId, token, "L");
        time.Advance(TimeSpan.FromMinutes(20));
        using (var refresh = await host.Post(fileId, token, "REFRESH_LOCK", ("X-WOPI-Lock", "L")))
        {
            await Assert.That(refresh.StatusCode).IsEqualTo(HttpStatusCode.OK);
            await Assert.That(Header(refresh, "X-WOPI-Lock")).IsNull();
        }

        time.Advance(TimeSpan.FromMinutes(20));
        await Assert.That(await CurrentLock(host, fileId, token)).IsEqualTo("L");
        time.Advance(TimeSpan.FromMinutes(11));
        await Assert.That(await CurrentLock(host, fileId, token)).IsEqualTo("");
    }

    /// <summary>
    /// 3.3.5.1.18: X-WOPI-Lock MUST be included with a 409 if no lock or a different lock exists.
    /// </summary>
    [Test]
    public async Task RefreshLock_Mismatch()
    {
        var (host, fileId, token) = await Open();
        await using var _ = host;
        using (var noLock = await host.Post(fileId, token, "REFRESH_LOCK", ("X-WOPI-Lock", "L")))
        {
            await Assert.That(noLock.StatusCode).IsEqualTo(HttpStatusCode.Conflict);
            await Assert.That(Header(noLock, "X-WOPI-Lock")).IsEqualTo("");
        }

        await Lock(host, fileId, token, "L");
        using var mismatch = await host.Post(fileId, token, "REFRESH_LOCK", ("X-WOPI-Lock", "other"));
        await Assert.That(mismatch.StatusCode).IsEqualTo(HttpStatusCode.Conflict);
        await Assert.That(Header(mismatch, "X-WOPI-Lock")).IsEqualTo("L");
    }

    // 3.3.5.1.22 Unlock

    /// <summary>
    /// 3.3.5.1.22: the lock is released. X-WOPI-Lock SHOULD NOT be included with a 200, and X-WOPI-ItemVersion SHOULD
    /// be the Version of CheckFileInfo.
    /// </summary>
    [Test]
    public async Task Unlock_MatchingLock()
    {
        var (host, fileId, token) = await Open();
        await using var _ = host;
        await Lock(host, fileId, token, "L");
        using var response = await host.Post(fileId, token, "UNLOCK", ("X-WOPI-Lock", "L"));
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.OK);
        await Assert.That(Header(response, "X-WOPI-Lock")).IsNull();
        await Assert.That(Header(response, "X-WOPI-ItemVersion")).IsEqualTo(await Version(host, fileId, token));
        await Assert.That(await CurrentLock(host, fileId, token)).IsEqualTo("");
    }

    /// <summary>
    /// 3.3.5.1.22: a lock mismatch is a 409, where X-WOPI-Lock SHOULD give the current lock; the lock is kept.
    /// </summary>
    [Test]
    public async Task Unlock_Mismatch()
    {
        var (host, fileId, token) = await Open();
        await using var _ = host;
        using (var noLock = await host.Post(fileId, token, "UNLOCK", ("X-WOPI-Lock", "L")))
        {
            await Assert.That(noLock.StatusCode).IsEqualTo(HttpStatusCode.Conflict);
            await Assert.That(Header(noLock, "X-WOPI-Lock")).IsEqualTo("");
        }

        await Lock(host, fileId, token, "L");
        using (var mismatch = await host.Post(fileId, token, "UNLOCK", ("X-WOPI-Lock", "other")))
        {
            await Assert.That(mismatch.StatusCode).IsEqualTo(HttpStatusCode.Conflict);
            await Assert.That(Header(mismatch, "X-WOPI-Lock")).IsEqualTo("L");
        }

        await Assert.That(await CurrentLock(host, fileId, token)).IsEqualTo("L");
    }

    [Test]
    public async Task Unlock_UnknownFile()
    {
        var (host, fileId, token) = await Open();
        await using var _ = host;
        await Lock(host, fileId, token, "L");
        File.Delete(host.PathOf("/test.docx"));
        using var response = await host.Post(fileId, token, "UNLOCK", ("X-WOPI-Lock", "L"));
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.NotFound);
    }

    // 3.3.5.1.23 UnlockAndRelock

    /// <summary>
    /// 3.3.5.1.23: X-WOPI-OldLock identifies the lock to release, X-WOPI-Lock is the new lock, which MUST expire
    /// 30 minutes after the operation.
    /// </summary>
    [Test]
    public async Task UnlockAndRelock_MatchingOldLock()
    {
        var time = new ManualTime();
        var (host, fileId, token) = await Open(time: time);
        await using var _ = host;
        await Lock(host, fileId, token, "L1");
        time.Advance(TimeSpan.FromMinutes(20));
        using (var response = await host.Post(fileId, token, "LOCK", ("X-WOPI-Lock", "L2"), ("X-WOPI-OldLock", "L1")))
        {
            await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.OK);
            await Assert.That(Header(response, "X-WOPI-Lock")).IsNull();
        }

        time.Advance(TimeSpan.FromMinutes(20));
        await Assert.That(await CurrentLock(host, fileId, token)).IsEqualTo("L2");
    }

    /// <summary>
    /// 3.3.5.1.23: the operation MUST be atomic: on a mismatch of the old lock, the current lock is kept and given
    /// in X-WOPI-Lock.
    /// </summary>
    [Test]
    public async Task UnlockAndRelock_Mismatch()
    {
        var (host, fileId, token) = await Open();
        await using var _ = host;
        await Lock(host, fileId, token, "L1");
        using var response = await host.Post(fileId, token, "LOCK", ("X-WOPI-Lock", "L2"), ("X-WOPI-OldLock", "other"));
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.Conflict);
        await Assert.That(Header(response, "X-WOPI-Lock")).IsEqualTo("L1");
        await Assert.That(await CurrentLock(host, fileId, token)).IsEqualTo("L1");
    }

    /// <summary>
    /// 3.3.5.1.23: there is no lock to release on an unlocked file.
    /// </summary>
    [Test]
    public async Task UnlockAndRelock_UnlockedFile()
    {
        var (host, fileId, token) = await Open();
        await using var _ = host;
        using var response = await host.Post(fileId, token, "LOCK", ("X-WOPI-Lock", "L2"), ("X-WOPI-OldLock", "L1"));
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.Conflict);
        await Assert.That(Header(response, "X-WOPI-Lock")).IsEqualTo("");
        await Assert.That(await CurrentLock(host, fileId, token)).IsEqualTo("");
    }

    // 3.3.5.1.15 PutRelativeFile

    static async Task<(HttpResponseMessage Response, JsonObject? Json)> PutRelative(TestServerHost host, string fileId, string token, string content, params (string Name, string Value)[] headers)
    {
        using var request = new HttpRequestMessage(HttpMethod.Post, $"/wopi/files/{fileId}?access_token={token}");
        request.Headers.Add("X-WOPI-Override", "PUT_RELATIVE");
        request.Headers.Add("X-WOPI-Size", Encoding.UTF8.GetByteCount(content).ToString());
        foreach (var (name, value) in headers)
        {
            request.Headers.TryAddWithoutValidation(name, value);
        }

        request.Content = new StringContent(content);
        var response = await host.Client.SendAsync(request);
        JsonObject? json = null;
        if (response.StatusCode == HttpStatusCode.OK)
        {
            json = JsonNode.Parse(await Body(response))!.AsObject();
        }

        return (response, json);
    }

    /// <summary>
    /// 3.3.5.1.15: with only an extension, the name of the initial file without extension SHOULD be combined with it.
    /// The server MUST use the content in the POST body to create the new file.
    /// </summary>
    [Test]
    public async Task PutRelativeFile_SuggestedExtension()
    {
        var (host, fileId, token) = await Open();
        await using var _ = host;
        var (response, json) = await PutRelative(host, fileId, token, "pdf content", ("X-WOPI-SuggestedTarget", ".pdf"));
        using (response)
        {
            await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.OK);
        }

        await Assert.That((string?) json!["Name"]).IsEqualTo("test.pdf");
        await Assert.That(await File.ReadAllTextAsync(host.PathOf("/test.pdf"))).IsEqualTo("pdf content");
    }

    /// <summary>
    /// 3.3.5.1.15: the server MUST modify the suggested name so the new file does not overwrite any existing one,
    /// while preserving the file extension.
    /// </summary>
    [Test]
    public async Task PutRelativeFile_SuggestedNameExists()
    {
        var (host, fileId, token) = await Open();
        await using var _ = host;
        await File.WriteAllTextAsync(host.PathOf("/report.docx"), "existing");
        var (response, json) = await PutRelative(host, fileId, token, "new", ("X-WOPI-SuggestedTarget", "report.docx"));
        using (response)
        {
            await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.OK);
        }

        var name = (string) json!["Name"]!;
        await Assert.That(name).IsNotEqualTo("report.docx");
        await Assert.That(name).EndsWith(".docx");
        await Assert.That(await File.ReadAllTextAsync(host.PathOf("/report.docx"))).IsEqualTo("existing");
        await Assert.That(await File.ReadAllTextAsync(host.PathOf($"/{name}"))).IsEqualTo("new");
    }

    /// <summary>
    /// 3.3.5.1.15: the server MUST modify the suggested name as needed to be legally named.
    /// </summary>
    [Test]
    public async Task PutRelativeFile_SuggestedIllegalName()
    {
        var (host, fileId, token) = await Open();
        await using var _ = host;
        var (response, json) = await PutRelative(host, fileId, token, "new", ("X-WOPI-SuggestedTarget", "../escape.docx"));
        using (response)
        {
            await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.OK);
        }

        var name = (string) json!["Name"]!;
        await Assert.That(name).DoesNotContain("/");
        await Assert.That(File.Exists(host.PathOf($"/{name}"))).IsTrue();
        await Assert.That(Directory.GetFiles(Path.GetDirectoryName(host.StoragePath)!, "escape.docx")).IsEmpty();
    }

    /// <summary>
    /// 3.3.5.1.15: the targets are UTF-7 strings.
    /// </summary>
    [Test]
    public async Task PutRelativeFile_Utf7Names()
    {
        var (host, fileId, token) = await Open();
        await using var _ = host;
        // "été.docx"
        var (response, json) = await PutRelative(host, fileId, token, "new", ("X-WOPI-RelativeTarget", "+AOk-t+AOk-.docx"));
        using (response)
        {
            await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.OK);
        }

        await Assert.That((string?) json!["Name"]).IsEqualTo("été.docx");
        await Assert.That(File.Exists(host.PathOf("/été.docx"))).IsTrue();
    }

    /// <summary>
    /// 3.3.5.1.15: the server MUST NOT modify a relative target.
    /// </summary>
    [Test]
    public async Task PutRelativeFile_RelativeTarget()
    {
        var (host, fileId, token) = await Open();
        await using var _ = host;
        var (response, json) = await PutRelative(host, fileId, token, "new", ("X-WOPI-RelativeTarget", "exact name.docx"));
        using (response)
        {
            await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.OK);
        }

        await Assert.That((string?) json!["Name"]).IsEqualTo("exact name.docx");
        await Assert.That(await File.ReadAllTextAsync(host.PathOf("/exact name.docx"))).IsEqualTo("new");
    }

    /// <summary>
    /// 3.3.5.1.15: an existing relative target without X-WOPI-OverwriteRelativeTarget is a 409, which SHOULD include
    /// X-WOPI-ValidRelativeTarget, a name usable in a further PutRelativeFile.
    /// </summary>
    [Test]
    public async Task PutRelativeFile_RelativeTargetExists()
    {
        var (host, fileId, token) = await Open();
        await using var _ = host;
        await File.WriteAllTextAsync(host.PathOf("/other.docx"), "existing");
        var validTarget = "";
        foreach (var overwrite in new[] { null, "false" })
        {
            (string, string)[] headers = overwrite == null ?
                [("X-WOPI-RelativeTarget", "other.docx")] :
                [("X-WOPI-RelativeTarget", "other.docx"), ("X-WOPI-OverwriteRelativeTarget", overwrite)];
            var (conflict, _) = await PutRelative(host, fileId, token, "new", headers);
            using (conflict)
            {
                await Assert.That(conflict.StatusCode).IsEqualTo(HttpStatusCode.Conflict);
                validTarget = Header(conflict, "X-WOPI-ValidRelativeTarget")!;
            }

            await Assert.That(validTarget).IsNotNull();
            await Assert.That(File.Exists(host.PathOf($"/{FileNames.DecodeUtf7(validTarget)}"))).IsFalse();
        }

        await Assert.That(await File.ReadAllTextAsync(host.PathOf("/other.docx"))).IsEqualTo("existing");
        var (retry, json) = await PutRelative(host, fileId, token, "new", ("X-WOPI-RelativeTarget", validTarget));
        using (retry)
        {
            await Assert.That(retry.StatusCode).IsEqualTo(HttpStatusCode.OK);
        }

        await Assert.That((string?) json!["Name"]).IsEqualTo(FileNames.DecodeUtf7(validTarget));
    }

    /// <summary>
    /// 3.3.5.1.15: with X-WOPI-OverwriteRelativeTarget, an existing target is overwritten unless it is locked.
    /// </summary>
    [Test]
    public async Task PutRelativeFile_OverwriteRelativeTarget()
    {
        await using var host = await TestServerHost.Start();
        await File.WriteAllTextAsync(host.PathOf("/test.docx"), databuf);
        await File.WriteAllTextAsync(host.PathOf("/other.docx"), "existing");
        var (fileId, token) = await host.Open("/test.docx");
        var (otherId, otherToken) = await host.Open("/other.docx");
        await Lock(host, otherId, otherToken, "L");

        var (locked, _) = await PutRelative(host, fileId, token, "new", ("X-WOPI-RelativeTarget", "other.docx"), ("X-WOPI-OverwriteRelativeTarget", "true"));
        using (locked)
        {
            await Assert.That(locked.StatusCode).IsEqualTo(HttpStatusCode.Conflict);
        }

        await Assert.That(await File.ReadAllTextAsync(host.PathOf("/other.docx"))).IsEqualTo("existing");
        using (var unlock = await host.Post(otherId, otherToken, "UNLOCK", ("X-WOPI-Lock", "L")))
        {
            await Assert.That(unlock.StatusCode).IsEqualTo(HttpStatusCode.OK);
        }

        var (overwrite, _) = await PutRelative(host, fileId, token, "new", ("X-WOPI-RelativeTarget", "other.docx"), ("X-WOPI-OverwriteRelativeTarget", "true"));
        using (overwrite)
        {
            await Assert.That(overwrite.StatusCode).IsEqualTo(HttpStatusCode.OK);
        }

        await Assert.That(await File.ReadAllTextAsync(host.PathOf("/other.docx"))).IsEqualTo("new");
    }

    /// <summary>
    /// 3.3.5.1.15: an illegal name is a 400; a relative target cannot be modified to make it legal.
    /// </summary>
    [Test]
    [Arguments("../escape.docx")]
    [Arguments("sub/file.docx")]
    [Arguments("..")]
    public async Task PutRelativeFile_RelativeTargetIllegal(string target)
    {
        var (host, fileId, token) = await Open();
        await using var _ = host;
        var (response, _) = await PutRelative(host, fileId, token, "new", ("X-WOPI-RelativeTarget", target));
        using (response)
        {
            await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.BadRequest);
        }
    }

    /// <summary>
    /// 3.3.5.1.15: X-WOPI-SuggestedTarget MUST be present if X-WOPI-RelativeTarget is not.
    /// </summary>
    [Test]
    public async Task PutRelativeFile_NoTarget()
    {
        var (host, fileId, token) = await Open();
        await using var _ = host;
        var (response, _) = await PutRelative(host, fileId, token, "new");
        using (response)
        {
            await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.BadRequest);
        }
    }

    /// <summary>
    /// 3.3.5.1.15.2: Name, Url, HostViewUrl and HostEditUrl are required. Url is
    /// HTTP://server/&lt;...&gt;/wopi*/files/&lt;id&gt;?access_token=&lt;token&gt; for the new file.
    /// </summary>
    [Test]
    public async Task PutRelativeFile_Response()
    {
        var (host, fileId, token) = await Open();
        await using var _ = host;
        var (response, json) = await PutRelative(host, fileId, token, "pdf content", ("X-WOPI-SuggestedTarget", ".pdf"));
        using (response)
        {
            await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.OK);
        }

        foreach (var name in new[] { "Name", "Url", "HostViewUrl", "HostEditUrl" })
        {
            await Assert.That(json![name]?.GetValueKind()).IsEqualTo(JsonValueKind.String);
        }

        var url = new Uri((string) json!["Url"]!);
        await Assert.That(Regex.IsMatch(url.AbsolutePath, "^/wopi/files/[^/]+$")).IsTrue();
        await Assert.That(url.Query).StartsWith("?access_token=");
        using var info = await host.Client.GetAsync(url.PathAndQuery);
        await Assert.That((string?) JsonNode.Parse(await Body(info))!["BaseFileName"]).IsEqualTo("test.pdf");
        using var contents = await host.Client.GetAsync($"{url.AbsolutePath}/contents{url.Query}");
        await Assert.That(await Body(contents)).IsEqualTo("pdf content");
    }

    /// <summary>
    /// 2.2.3: the token determines the permissions, so a read-only token cannot create files.
    /// </summary>
    [Test]
    public async Task PutRelativeFile_ReadOnlyToken()
    {
        var (host, fileId, token) = await Open(mode: Mode.ViewOnly);
        await using var _ = host;
        var (response, _) = await PutRelative(host, fileId, token, "new", ("X-WOPI-SuggestedTarget", ".pdf"));
        using (response)
        {
            await Assert.That(response.IsSuccessStatusCode).IsFalse();
        }

        await Assert.That(File.Exists(host.PathOf("/test.pdf"))).IsFalse();
    }

    // 3.3.5.1.19 RenameFile

    /// <summary>
    /// 3.3.5.1.19: the response Name is the new name without the path or file extension.
    /// </summary>
    [Test]
    public async Task RenameFile_WithLock()
    {
        var (host, fileId, token) = await Open();
        await using var _ = host;
        await Assert.That((bool?) (await CheckFileInfo(host, fileId, token))["SupportsRename"]).IsTrue();
        await Lock(host, fileId, token, "L");
        using var response = await host.Post(fileId, token, "RENAME_FILE", ("X-WOPI-Lock", "L"), ("X-WOPI-RequestedName", "renamed"));
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.OK);
        await Assert.That(Header(response, "X-WOPI-Lock")).IsNull();
        await Assert.That((string?) JsonNode.Parse(await Body(response))!["Name"]).IsEqualTo("renamed");
        await Assert.That(await File.ReadAllTextAsync(host.PathOf("/renamed.docx"))).IsEqualTo(databuf);
        await Assert.That(File.Exists(host.PathOf("/test.docx"))).IsFalse();
    }

    /// <summary>
    /// 3.3.5.1.19: X-WOPI-Lock MUST be included with a 409.
    /// </summary>
    [Test]
    public async Task RenameFile_LockMismatch()
    {
        var (host, fileId, token) = await Open();
        await using var _ = host;
        await Lock(host, fileId, token, "L");
        using var response = await host.Post(fileId, token, "RENAME_FILE", ("X-WOPI-Lock", "other"), ("X-WOPI-RequestedName", "renamed"));
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.Conflict);
        await Assert.That(Header(response, "X-WOPI-Lock")).IsEqualTo("L");
        await Assert.That(File.Exists(host.PathOf("/test.docx"))).IsTrue();
    }

    /// <summary>
    /// 3.3.5.1.19: an illegal name is a 400, which SHOULD include X-WOPI-InvalidFileNameError. A name that
    /// would replace another file is not usable either.
    /// </summary>
    [Test]
    [Arguments("../escape")]
    [Arguments("sub/renamed")]
    [Arguments("other")]
    public async Task RenameFile_IllegalName(string name)
    {
        var (host, fileId, token) = await Open();
        await using var _ = host;
        await File.WriteAllTextAsync(host.PathOf("/other.docx"), "existing");
        using var response = await host.Post(fileId, token, "RENAME_FILE", ("X-WOPI-Lock", ""), ("X-WOPI-RequestedName", name));
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.BadRequest);
        await Assert.That(string.IsNullOrEmpty(Header(response, "X-WOPI-InvalidFileNameError"))).IsFalse();
        await Assert.That(File.Exists(host.PathOf("/test.docx"))).IsTrue();
        await Assert.That(await File.ReadAllTextAsync(host.PathOf("/other.docx"))).IsEqualTo("existing");
    }

    /// <summary>
    /// 3.3.5.1.2.2: UserCanRename is false for a user who cannot rename, and the request is refused.
    /// </summary>
    [Test]
    public async Task RenameFile_NotAllowed()
    {
        var (host, fileId, _) = await Open();
        await using var _ = host;
        var token = await host.Token(TestServerHost.Access(fileId) with { CanRename = false });
        await Assert.That((bool?) (await CheckFileInfo(host, fileId, token))["UserCanRename"]).IsFalse();
        using var response = await host.Post(fileId, token, "RENAME_FILE", ("X-WOPI-Lock", ""), ("X-WOPI-RequestedName", "renamed"));
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.Unauthorized);
        await Assert.That(File.Exists(host.PathOf("/test.docx"))).IsTrue();
    }

    /// <summary>
    /// 2.2.3: the id identifies the file, so the WOPI session goes on with the same id and token after a rename.
    /// </summary>
    [Test]
    public async Task RenameFile_SessionContinues()
    {
        var (host, fileId, token) = await Open();
        await using var _ = host;
        using var response = await host.Post(fileId, token, "RENAME_FILE", ("X-WOPI-Lock", ""), ("X-WOPI-RequestedName", "renamed"));
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.OK);
        await Assert.That((string?) (await CheckFileInfo(host, fileId, token))["BaseFileName"]).IsEqualTo("renamed.docx");
    }

    // 3.3.5.1.5 DeleteFile

    /// <summary>
    /// 3.3.5.1.5: the server MUST delete the file if possible.
    /// </summary>
    [Test]
    public async Task DeleteFile_Unlocked()
    {
        var (host, fileId, token) = await Open();
        await using var _ = host;
        using (var response = await host.Post(fileId, token, "DELETE"))
        {
            await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.OK);
        }

        await Assert.That(File.Exists(host.PathOf("/test.docx"))).IsFalse();
        using var info = await host.Client.GetAsync($"/wopi/files/{fileId}?access_token={token}");
        await Assert.That(info.StatusCode).IsEqualTo(HttpStatusCode.NotFound);
    }

    /// <summary>
    /// 3.3.5.1.5: a locked file is a 409, where X-WOPI-Lock SHOULD give the lock.
    /// </summary>
    [Test]
    public async Task DeleteFile_Locked()
    {
        var (host, fileId, token) = await Open();
        await using var _ = host;
        await Lock(host, fileId, token, "L");
        using var response = await host.Post(fileId, token, "DELETE");
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.Conflict);
        await Assert.That(Header(response, "X-WOPI-Lock")).IsEqualTo("L");
        await Assert.That(File.Exists(host.PathOf("/test.docx"))).IsTrue();
    }

    [Test]
    public async Task DeleteFile_UnknownFile()
    {
        var (host, fileId, token) = await Open();
        await using var _ = host;
        File.Delete(host.PathOf("/test.docx"));
        using var response = await host.Post(fileId, token, "DELETE");
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.NotFound);
    }

    // 3.3.5.1 operations the server does not implement

    /// <summary>
    /// 3.3.5.1: the operations the server does not implement are answered with 501 (Unsupported), and CheckFileInfo
    /// does not advertise them.
    /// </summary>
    [Test]
    [Arguments("ADD_ACTIVITIES", "SupportsAddActivities")]
    [Arguments("CHECK_POLICY", "SupportsCheckPolicy")]
    [Arguments("CHECK_USER_ACCESS", "SupportsCheckUserAccess")]
    [Arguments("COBALT", "SupportsCobalt")]
    [Arguments("GET_ACTIVITIES", "SupportsGetActivities")]
    [Arguments("GET_FILE_USER_VALUE", "SupportsFileUserValue")]
    [Arguments("SET_FILE_USER_VALUE", "SupportsFileUserValue")]
    [Arguments("GET_RESTRICTED_LINK", "SupportsScenarioLinks")]
    [Arguments("REVOKE_RESTRICTED_LINK", "SupportsScenarioLinks")]
    [Arguments("GET_SHARE_URL", "SupportedShareUrlTypes")]
    [Arguments("GRANT_USER_ACCESS", "SupportsGrantUserAccess")]
    [Arguments("READ_SECURE_STORE", "SupportsSecureStore")]
    public async Task Unsupported_Operations(string operation, string property)
    {
        var (host, fileId, token) = await Open();
        await using var _ = host;
        var advertised = (await CheckFileInfo(host, fileId, token))[property];
        await Assert.That(advertised is null || advertised.GetValueKind() == JsonValueKind.False).IsTrue();
        using var response = await host.Post(fileId, token, operation);
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.NotImplemented);
    }

    /// <summary>
    /// 3.3.5.2 and 3.3.5.4: the server has no folders, so every folder is unknown (404) or the operations unsupported (501).
    /// SupportsFolders is not advertised.
    /// </summary>
    [Test]
    public async Task Unsupported_Folders()
    {
        var (host, fileId, token) = await Open();
        await using var _ = host;
        await Assert.That((await CheckFileInfo(host, fileId, token)).ContainsKey("SupportsFolders")).IsFalse();
        foreach (var path in new[] { $"/wopi/folders/{fileId}?access_token={token}", $"/wopi/folders/{fileId}/children?access_token={token}" })
        {
            using var response = await host.Client.GetAsync(path);
            await Assert.That(response.StatusCode is HttpStatusCode.NotFound or HttpStatusCode.NotImplemented).IsTrue();
        }
    }
}
