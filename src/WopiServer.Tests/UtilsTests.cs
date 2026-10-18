public class UtilsTests
{
    [Test]
    [Arguments("/a/b.docx", "/a", "b.docx", "/a/b", ".docx")]
    [Arguments("/b.docx", "/", "b.docx", "/b", ".docx")]
    [Arguments("b.docx", "", "b.docx", "b", ".docx")]
    [Arguments("/a/.bashrc", "/a", ".bashrc", "/a/.bashrc", "")]
    [Arguments("/a.b/c", "/a.b", "c", "/a.b/c", "")]
    [Arguments("/a/.~lock.x.odt#", "/a", ".~lock.x.odt#", "/a/.~lock.x", ".odt#")]
    [Arguments("//a//b", "//a", "b", "//a//b", "")]
    public async Task PosixPaths(string path, string dir, string name, string root, string extension)
    {
        await Assert.That(PosixPath.DirName(path)).IsEqualTo(dir);
        await Assert.That(PosixPath.BaseName(path)).IsEqualTo(name);
        await Assert.That(PosixPath.SplitExt(path)).IsEqualTo((root, extension));
    }

    [Test]
    public async Task QuotePlus()
    {
        await Assert.That(UrlCodec.QuotePlus("http://host:8880/wopi/files/a!b c~d(é)")).IsEqualTo("http%3A%2F%2Fhost%3A8880%2Fwopi%2Ffiles%2Fa%21b+c~d%28%C3%A9%29");
        await Assert.That(UrlCodec.UnquotePlus("http%3A%2F%2Fhost+x%2F%C3%A9%zz")).IsEqualTo("http://host x/é%zz");
        await Assert.That(UrlCodec.Unquote("a+b%20c")).IsEqualTo("a+b c");
    }

    [Test]
    [Arguments("REFRESH_LOCK", "Refresh_Lock")]
    [Arguments("CheckFileInfo", "Checkfileinfo")]
    [Arguments("genAccessToken", "Genaccesstoken")]
    [Arguments("a1b", "A1B")]
    public async Task Title(string value, string expected) =>
        await Assert.That(value.Title()).IsEqualTo(expected);

    [Test]
    [Arguments("../../etc/passwd", "etc_passwd")]
    [Arguments("My cool movie.mov", "My_cool_movie.mov")]
    [Arguments("i contain cool ümläuts.txt", "i_contain_cool_umlauts.txt")]
    [Arguments("root!0:0", "root00")]
    public async Task SecureFilename(string value, string expected) =>
        await Assert.That(PyStrings.SecureFilename(value)).IsEqualTo(expected);

    [Test]
    public async Task Utf7() =>
        await Assert.That(PyStrings.DecodeUtf7("Hi Mom -+Jjo--!")).IsEqualTo("Hi Mom -☺-!");

    [Test]
    public async Task Short()
    {
        await Assert.That("0123456789012345678901234".Short()).IsEqualTo("56789012345678901234");
        await Assert.That("abc".Short()).IsEqualTo("abc");
        await Assert.That(((string?) null).Short()).IsEqualTo("");
    }

    [Test]
    public async Task Inode_RoundTrip()
    {
        var inode = Common.EncodeInode("local", "123456");
        await Assert.That(inode).IsEqualTo("local!MTIzNDU2");
        await Assert.That(Common.DecodeInode(inode)).IsEqualTo(("local", "123456"));
        await Assert.That(Common.DecodeInode(Common.EncodeInode("eoshome", "ab?"))).IsEqualTo(("eoshome", "ab?"));
    }

    [Test]
    public async Task RevaLock_RoundTrip()
    {
        var raw = Common.GenRevaLock("app", "id", 1665446400);
        var json = Encoding.UTF8.GetString(Convert.FromBase64String(raw.Replace('-', '+').Replace('_', '/')));
        await Assert.That(json).IsEqualTo("""{"lock_id":"id","type":2,"app_name":"app","user":{},"expiration":{"seconds":1665446400}}""");
        var storageLock = Common.RetrieveRevaLock(raw.TrimEnd('='));
        await Assert.That(storageLock.LockId).IsEqualTo("id");
        await Assert.That(storageLock.AppName).IsEqualTo("app");
        await Assert.That(storageLock.Expiration.Seconds).IsEqualTo(1665446400);
        var exception = await Helpers.Throws<StorageException>(() => Task.FromResult(Common.RetrieveRevaLock("!!")));
        await Assert.That(exception.Message).StartsWith("Unable to parse existing lock");
    }

    [Test]
    public async Task EncodeLock()
    {
        await Assert.That(WopiUtils.EncodeLock("abc")).IsEqualTo($"{Common.WebDavLockPrefix} YWJj");
        await Assert.That(WopiUtils.EncodeLock(null)).IsNull();
        await Assert.That(WopiUtils.EncodeLock("")).IsNull();
    }

    [Test]
    [Arguments("/a/test.docx", "/a/~$test.docx")]
    [Arguments("/a/abcdefgh.docx", "/a/~$cdefgh.docx")]
    [Arguments("/a/abcdefg.docx", "/a/~$bcdefg.docx")]
    [Arguments("/a/abcdefgh.xlsx", "/a/~$abcdefgh.xlsx")]
    public async Task MicrosoftOfficeLockName(string filename, string expected) =>
        await Assert.That(WopiUtils.GetMicrosoftOfficeLockName(filename)).IsEqualTo(expected);

    [Test]
    public async Task LibreOfficeLockName() =>
        await Assert.That(WopiUtils.GetLibreOfficeLockName("/a/test.odt")).IsEqualTo("/a/.~lock.test.odt#");

    [Test]
    public async Task Jwt_RoundTrip()
    {
        var token = Jwt.Encode(new JsonObject { ["u"] = "x", ["exp"] = 2000 }, "short");
        await Assert.That(token).StartsWith("eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.");
        var payload = Jwt.Decode<JsonObject>(token, "short", 1000);
        await Assert.That((string?) payload["u"]).IsEqualTo("x");

        var expired = await Helpers.Throws<JwtException>(() => Task.FromResult(Jwt.Decode<JsonObject>(token, "short", 2000)));
        await Assert.That(expired.Expired).IsTrue();
        var invalid = await Helpers.Throws<JwtException>(() => Task.FromResult(Jwt.Decode<JsonObject>(token, "other", 1000)));
        await Assert.That(invalid.Message).IsEqualTo("Signature verification failed");
        await Helpers.Throws<JwtException>(() => Task.FromResult(Jwt.Decode<JsonObject>("abc", "short", 1000)));
    }

    /// <summary>
    /// The token PyJWT 2 issues for <c>jwt.encode({"u": "x", "exp": 2000}, "short", algorithm="HS256")</c>: compact JSON, sorted header. The signature was computed independently with <c>openssl dgst -sha256 -hmac short</c>.
    /// </summary>
    [Test]
    public async Task Jwt_PyJwtToken()
    {
        var token = Jwt.Encode(Encoding.UTF8.GetBytes("""{"u":"x","exp":2000}"""), "short");
        await Assert.That(token).IsEqualTo("eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJ1IjoieCIsImV4cCI6MjAwMH0.ywacz_9Tn_HB-3L1EQhuUE438K3YWaRQPIjXhpufYtk");
    }

    [Test]
    public async Task Jsonify()
    {
        await Assert.That(WopiClient.Jsonify("File saved successfully")).IsEqualTo("""{"message": "File saved successfully", "delay": "1.6"}""");
        await Assert.That(WopiClient.Jsonify(new('x', 61))).EndsWith("\", \"delay\": \"0.0\"}");
        await Assert.That(WopiClient.Jsonify("a \"quoted\" one")).IsEqualTo("""{"message": "a \"quoted\" one", "delay": "1.2"}""");
    }

    static WopiUtils Utils(bool strict)
    {
        var config = Helpers.Ini(
            $"""
             [general]
             wopilockstrictcheck = {strict}
             """);
        var state = new WopiState
        {
            Config = new(config),
            Storage = null!,
            Time = TimeProvider.System,
            StorageType = "local",
            Port = 0,
            UseHttps = false,
            WopiSecret = "",
            IopSecret = "",
            WopiUrl = "http://localhost",
            HomePath = "",
            RecoveryPath = "",
            WopiProxy = "",
            WopiProxyKey = null,
            ProxiedAppName = "",
            NonOfficeTypes = [],
            CodeTypes = [],
            EnableRename = false
        };
        return new(state, NullLogger.Instance);
    }

    [Test]
    [Arguments("abc", "abc", true)]
    [Arguments("abc", "abd", false)]
    [Arguments("""{"S":"1","F":2}""", """{"S":"1","F":3}""", false)]
    public async Task CompareWopiLocks_Strict(string lock1, string lock2, bool expected) =>
        await Assert.That(Utils(true).CompareWopiLocks(lock1, lock2)).IsEqualTo(expected);

    [Test]
    [Arguments("""{"S":"1","F":2}""", """{"S":"1","F":3}""", true)]
    [Arguments("""{"S":"1","F":2}""", """{"S":"2","F":2}""", false)]
    [Arguments("""{"S":"1","F":2}""", "1", true)]
    [Arguments("""{"S":"1","F":2}""", """{"F":2}""", false)]
    [Arguments("""{"F":2}""", "1", false)]
    [Arguments("abc", "abd", false)]
    public async Task CompareWopiLocks_Relaxed(string lock1, string lock2, bool expected) =>
        await Assert.That(Utils(false).CompareWopiLocks(lock1, lock2)).IsEqualTo(expected);

    [Test]
    public async Task UrlTemplate()
    {
        const string template = "https://efss/external-<app><path>?id=<endpoint>!<fileid>";
        var regular = WopiUtils.GenerateUrlFromTemplate(template, "Alice", UserType.Regular, "https://efss/files/x", "/home/a.docx", "storage", "123", "Office");
        await Assert.That(regular).IsEqualTo("https://efss/external-Office/home/a.docx?id=storage!123");
        var anonymous = WopiUtils.GenerateUrlFromTemplate(template, "", UserType.Anonymous, "https://efss/files/link/public/abc?x=1", "/home/a.docx", "storage", "123", "Office");
        await Assert.That(anonymous).IsEqualTo("https://efss/external-Office/public/abc/a.docx?id=storage!123");
    }
}
