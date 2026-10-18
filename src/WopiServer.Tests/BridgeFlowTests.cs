using Microsoft.AspNetCore.WebUtilities;

/// <summary>
/// The bridge end to end against stubbed apps: open, save loop, failures and recovery.
/// </summary>
public class BridgeFlowTests
{
    const string attachment = "upload_0123456789abcdef0123456789abcdef.png";

    static string? Header(HttpResponseMessage response, string name)
    {
        if (response.Headers.TryGetValues(name, out var values))
        {
            return string.Join(", ", values);
        }

        return null;
    }

    /// <summary>
    /// A document open through the bridge, as the app sees it.
    /// </summary>
    sealed record Opened(string AppUrl, string DocId, string WopiSrc, string Token);

    sealed class Context(TestServerHost host, ManualTime time) : IAsyncDisposable
    {
        public TestServerHost Host { get; } = host;
        public ManualTime Time { get; } = time;
        public Bridge Bridge => Host.Service<Bridge>();

        public async Task<HttpResponseMessage> Open(string file, string appName = "CodiMD", string appUrl = "http://codimd", string viewMode = "VIEW_MODE_READ_WRITE", string? apiKey = null) =>
            await Host.OpenInApp(file, viewMode, appName, appUrl, apiKey: apiKey);

        public async Task<Opened> OpenCodiMD(string file, string viewMode = "VIEW_MODE_READ_WRITE")
        {
            using var response = await Open(file, viewMode: viewMode);
            var appUrl = await AppUrl(response);
            var uri = new Uri(appUrl);
            var query = QueryHelpers.ParseQuery(uri.Query);
            return new(appUrl, uri.AbsolutePath.Trim('/'), query["wopiSrc"].ToString(), query["accessToken"].ToString());
        }

        public static async Task<string> AppUrl(HttpResponseMessage response)
        {
            var body = await response.Content.ReadAsStringAsync();
            if (response.StatusCode != HttpStatusCode.OK)
            {
                throw new($"openinapp failed: {response.StatusCode} {body}");
            }

            return (string) JsonNode.Parse(body)!["app-url"]!;
        }

        /// <summary>
        /// The save webhook the app calls when the document changes or is closed.
        /// </summary>
        public async Task<HttpResponseMessage> Save(Opened opened, bool close, string app = "CodiMD")
        {
            using var request = new HttpRequestMessage(HttpMethod.Post, $"/wopi/bridge/{opened.DocId}?WOPISrc={Uri.EscapeDataString(opened.WopiSrc)}&access_token={opened.Token}&close={(close ? "true" : "false")}");
            request.Headers.Add("X-Efss-Bridged-App", app);
            return await Host.Client.SendAsync(request);
        }

        public async Task<JsonObject> OpenFiles()
        {
            using var response = await Host.Client.GetAsync($"/wopi/bridge/list?apikey={TestServerHost.WopiSecret}");
            return JsonNode.Parse(await response.Content.ReadAsStringAsync())!.AsObject();
        }

        public async Task<string?> WopiLock(string file)
        {
            var (fileId, token, _) = await Host.Open(file);
            using var response = await Host.Post(fileId, token, "GET_LOCK");
            return Header(response, "X-WOPI-Lock");
        }

        public ValueTask DisposeAsync() =>
            Host.DisposeAsync();
    }

    static async Task<Context> Start(Func<HttpRequestMessage, Task<HttpResponseMessage>> apps, string extraConfig = "")
    {
        var time = new ManualTime();
        var host = await TestServerHost.Start(extraConfig, time, apps: apps);
        return new(host, time);
    }

    // CodiMD

    /// <summary>
    /// A read-only open pushes the note to a new, locked, CodiMD note, shown in its publish view.
    /// </summary>
    [Test]
    public async Task CodiMD_ReadOnly()
    {
        var codimd = new CodiMDStub();
        await using var context = await Start(codimd.Handle);
        await File.WriteAllTextAsync(context.Host.PathOf("/notes.md"), "# Read only");
        using var response = await context.Open("/notes.md", viewMode: "VIEW_MODE_READ_ONLY");
        await Assert.That(await Context.AppUrl(response)).IsEqualTo("http://codimd/new0/publish");
        await Assert.That(codimd.Notes["new0"]).IsEqualTo("# Read only");
        await Assert.That(await context.WopiLock("/notes.md")).IsEqualTo("");
        await Assert.That((await context.OpenFiles()).Count).IsEqualTo(0);
    }

    /// <summary>
    /// A docid aliased by CodiMD is followed.
    /// </summary>
    [Test]
    public async Task CodiMD_AliasedDocId()
    {
        var codimd = new CodiMDStub();
        await using var context = await Start(codimd.Handle);
        await File.WriteAllTextAsync(context.Host.PathOf("/notes.md"), "# Aliased");
        // the docid is a hash of the WOPISrc: learn it from a first open
        var opened = await context.OpenCodiMD("/notes.md");
        codimd.Notes.Clear();
        codimd.Alias = (opened.DocId, "alias");
        await using var other = await Start(codimd.Handle);
        await File.WriteAllTextAsync(other.Host.PathOf("/notes.md"), "# Aliased");
        var aliased = await other.OpenCodiMD("/notes.md");
        await Assert.That(aliased.DocId).IsEqualTo("alias");
        await Assert.That(codimd.Notes["alias"]).IsEqualTo("# Aliased");
    }

    /// <summary>
    /// Closing an unchanged note does not save it.
    /// </summary>
    [Test]
    public async Task CodiMD_CloseUnchanged()
    {
        var codimd = new CodiMDStub();
        await using var context = await Start(codimd.Handle);
        await File.WriteAllTextAsync(context.Host.PathOf("/notes.md"), "# Unchanged\n");
        var written = File.GetLastWriteTimeUtc(context.Host.PathOf("/notes.md"));
        var opened = await context.OpenCodiMD("/notes.md");
        using (var save = await context.Save(opened, close: true))
        {
            await Assert.That(save.StatusCode).IsEqualTo(HttpStatusCode.Accepted);
        }

        await context.Bridge.RunRound();
        using (var again = await context.Save(opened, close: true))
        {
            await Assert.That(again.StatusCode).IsEqualTo(HttpStatusCode.Accepted);
        }

        await Assert.That(File.GetLastWriteTimeUtc(context.Host.PathOf("/notes.md"))).IsEqualTo(written);
    }

    /// <summary>
    /// A .zmd bundle: on open its pictures are uploaded to CodiMD, on save they are bundled back with the note.
    /// </summary>
    [Test]
    public async Task CodiMD_Bundle()
    {
        var codimd = new CodiMDStub();
        await using var context = await Start(codimd.Handle);
        var markdown = $"# Bundle\n![picture](/byoa/codimd/uploads/{attachment})\n";
        await File.WriteAllBytesAsync(context.Host.PathOf("/notes.zmd"), Bundle.Create(("notes.md", Encoding.UTF8.GetBytes(markdown)), (attachment, [1, 2, 3])));
        var opened = await context.OpenCodiMD("/notes.zmd");
        await Assert.That(codimd.Uploads[attachment]).IsEquivalentTo(new byte[] { 1, 2, 3 });
        // the legacy reverse proxy path is dropped
        await Assert.That(codimd.Notes[opened.DocId]).IsEqualTo($"# Bundle\n![picture](/uploads/{attachment})\n");

        codimd.Notes[opened.DocId] += "Edited\n";
        using (var save = await context.Save(opened, close: true))
        {
            await Assert.That(save.StatusCode).IsEqualTo(HttpStatusCode.Accepted);
        }

        await context.Bridge.RunRound();
        var saved = Bundle.Read(await File.ReadAllBytesAsync(context.Host.PathOf("/notes.zmd")));
        await Assert.That(saved["notes.md"]).EndsWith("Edited\n");
        await Assert.That(saved.ContainsKey(attachment)).IsTrue();
    }

    /// <summary>
    /// A bundle whose pictures already exist in CodiMD with the same size does not upload them again.
    /// </summary>
    [Test]
    public async Task CodiMD_BundleExistingAttachment()
    {
        var codimd = new CodiMDStub();
        codimd.Uploads[attachment] = [1, 2, 3];
        await using var context = await Start(codimd.Handle);
        await File.WriteAllBytesAsync(context.Host.PathOf("/notes.zmd"), Bundle.Create(("notes.md", "# Bundle"u8.ToArray()), (attachment, [4, 5, 6])));
        await context.OpenCodiMD("/notes.zmd");
        await Assert.That(codimd.Uploads[attachment]).IsEquivalentTo(new byte[] { 1, 2, 3 });
    }

    [Test]
    public async Task CodiMD_InvalidBundle()
    {
        var codimd = new CodiMDStub();
        await using var context = await Start(codimd.Handle);
        await File.WriteAllTextAsync(context.Host.PathOf("/notes.zmd"), "not a zip");
        using var response = await context.Open("/notes.zmd");
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.InternalServerError);
        await Assert.That(await response.Content.ReadAsStringAsync()).IsEqualTo("The file is not in the expected zipped format");
    }

    [Test]
    public async Task CodiMD_InvalidUtf8()
    {
        var codimd = new CodiMDStub();
        await using var context = await Start(codimd.Handle);
        await File.WriteAllBytesAsync(context.Host.PathOf("/notes.md"), [0x23, 0x20, 0xC3, 0x28]);
        using var response = await context.Open("/notes.md");
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.InternalServerError);
        await Assert.That(await response.Content.ReadAsStringAsync()).StartsWith("File contains an invalid UTF-8 character");
    }

    /// <summary>
    /// A plain note given pictures is saved as a .zmd bundle on close, replacing the .md file.
    /// </summary>
    [Test]
    public async Task CodiMD_NewAttachmentsSaveAsBundle()
    {
        var codimd = new CodiMDStub();
        await using var context = await Start(codimd.Handle);
        await File.WriteAllTextAsync(context.Host.PathOf("/notes.md"), "# Plain\n");
        var opened = await context.OpenCodiMD("/notes.md");
        codimd.Uploads[attachment] = [1, 2, 3];
        codimd.Notes[opened.DocId] += $"![picture](/uploads/{attachment})\n";
        using (var save = await context.Save(opened, close: true))
        {
            await Assert.That(save.StatusCode).IsEqualTo(HttpStatusCode.Accepted);
        }

        await context.Bridge.RunRound();
        await Assert.That(File.Exists(context.Host.PathOf("/notes.md"))).IsFalse();
        var saved = Bundle.Read(await File.ReadAllBytesAsync(context.Host.PathOf("/notes.zmd")));
        await Assert.That(saved.Keys.Order()).IsEquivalentTo(new[] { "notes.md", attachment });
    }

    /// <summary>
    /// With disablezip, a plain note stays a plain note, whatever pictures it references.
    /// </summary>
    [Test]
    public async Task CodiMD_DisableZip()
    {
        var codimd = new CodiMDStub();
        await using var context = await Start(codimd.Handle, "[bridge]\ndisablezip = True\n[general]");
        await File.WriteAllTextAsync(context.Host.PathOf("/notes.md"), "# Plain\n");
        var opened = await context.OpenCodiMD("/notes.md");
        codimd.Notes[opened.DocId] += $"![picture](/uploads/{attachment})\n";
        using (var save = await context.Save(opened, close: true))
        {
            await Assert.That(save.StatusCode).IsEqualTo(HttpStatusCode.Accepted);
        }

        await context.Bridge.RunRound();
        await Assert.That(await File.ReadAllTextAsync(context.Host.PathOf("/notes.md"))).EndsWith($"/uploads/{attachment})\n");
        await Assert.That(File.Exists(context.Host.PathOf("/notes.zmd"))).IsFalse();
    }

    /// <summary>
    /// A picture that cannot be fetched from CodiMD is reported, and the note saved without it.
    /// </summary>
    [Test]
    public async Task CodiMD_MissingAttachment()
    {
        var codimd = new CodiMDStub();
        await using var context = await Start(codimd.Handle);
        await File.WriteAllBytesAsync(context.Host.PathOf("/notes.zmd"), Bundle.Create(("notes.md", "# Bundle\n"u8.ToArray())));
        var opened = await context.OpenCodiMD("/notes.zmd");
        codimd.Notes[opened.DocId] += $"![picture](/uploads/{attachment})\n";
        using (var save = await context.Save(opened, close: false))
        {
            await Assert.That(save.StatusCode).IsEqualTo(HttpStatusCode.Accepted);
        }

        context.Time.Advance(TimeSpan.FromSeconds(201));
        await context.Bridge.RunRound();
        using var again = await context.Save(opened, close: false);
        await Assert.That(again.StatusCode).IsEqualTo(HttpStatusCode.NotFound);
        await Assert.That(await again.Content.ReadAsStringAsync()).Contains("Failed to include a referenced picture in the saved file");
        await Assert.That(Bundle.Read(await File.ReadAllBytesAsync(context.Host.PathOf("/notes.zmd")))["notes.md"]).EndsWith($"{attachment})\n");
    }

    /// <summary>
    /// A note CodiMD fails to give back is retried after the save interval, and the file is not unlocked meanwhile.
    /// </summary>
    [Test]
    public async Task CodiMD_FetchFailsThenRetries()
    {
        var codimd = new CodiMDStub();
        await using var context = await Start(codimd.Handle);
        await File.WriteAllTextAsync(context.Host.PathOf("/notes.md"), "# Retry\n");
        var opened = await context.OpenCodiMD("/notes.md");
        codimd.Notes[opened.DocId] = "# Retry\nEdited\n";
        codimd.DownloadFails = true;
        using (var save = await context.Save(opened, close: true))
        {
            await Assert.That(save.StatusCode).IsEqualTo(HttpStatusCode.Accepted);
        }

        await context.Bridge.RunRound();
        using (var failed = await context.Save(opened, close: true))
        {
            await Assert.That(failed.StatusCode).IsEqualTo(HttpStatusCode.FailedDependency);
        }

        await Assert.That(await context.WopiLock("/notes.md")).Contains("still-dirty");
        codimd.DownloadFails = false;
        // a retry waits for the save interval
        await context.Bridge.RunRound();
        await Assert.That(await File.ReadAllTextAsync(context.Host.PathOf("/notes.md"))).IsEqualTo("# Retry\n");
        context.Time.Advance(TimeSpan.FromSeconds(201));
        await context.Bridge.RunRound();
        await Assert.That(await File.ReadAllTextAsync(context.Host.PathOf("/notes.md"))).IsEqualTo("# Retry\nEdited\n");
        await Assert.That(await context.WopiLock("/notes.md")).DoesNotContain("still-dirty");
    }

    /// <summary>
    /// A document left open with no activity is closed and unlocked by the save loop.
    /// </summary>
    [Test]
    public async Task CodiMD_IdleDocumentIsUnlocked()
    {
        var codimd = new CodiMDStub();
        await using var context = await Start(codimd.Handle);
        await File.WriteAllTextAsync(context.Host.PathOf("/notes.md"), "# Idle\n");
        await context.OpenCodiMD("/notes.md");
        await Assert.That(await context.WopiLock("/notes.md")).IsNotEqualTo("");
        await context.Bridge.RunRound();
        await Assert.That((await context.OpenFiles()).Count).IsEqualTo(1);

        context.Time.Advance(TimeSpan.FromMinutes(15));
        await context.Bridge.RunRound();
        await Assert.That(await context.WopiLock("/notes.md")).IsEqualTo("");
        await Assert.That((await context.OpenFiles()).Count).IsEqualTo(0);
    }

    /// <summary>
    /// A closed document is unlocked once the unlock interval passed, so a quick reopen finds it still locked.
    /// </summary>
    [Test]
    public async Task CodiMD_ClosedDocumentIsUnlockedAfterInterval()
    {
        var codimd = new CodiMDStub();
        await using var context = await Start(codimd.Handle);
        await File.WriteAllTextAsync(context.Host.PathOf("/notes.md"), "# Close\n");
        var opened = await context.OpenCodiMD("/notes.md");
        codimd.Notes[opened.DocId] = "# Close\nEdited\n";
        using (var save = await context.Save(opened, close: true))
        {
            await Assert.That(save.StatusCode).IsEqualTo(HttpStatusCode.Accepted);
        }

        await context.Bridge.RunRound();
        await Assert.That(await context.WopiLock("/notes.md")).IsNotEqualTo("");
        context.Time.Advance(TimeSpan.FromSeconds(91));
        await context.Bridge.RunRound();
        await Assert.That(await context.WopiLock("/notes.md")).IsEqualTo("");
        await Assert.That((await context.OpenFiles()).Count).IsEqualTo(0);
    }

    /// <summary>
    /// A second user joining the document is added to the lock, and the document stays locked when only one closes.
    /// </summary>
    [Test]
    public async Task CodiMD_TwoSessions()
    {
        var codimd = new CodiMDStub();
        await using var context = await Start(codimd.Handle);
        await File.WriteAllTextAsync(context.Host.PathOf("/notes.md"), "# Shared\n");
        var first = await context.OpenCodiMD("/notes.md");
        // another token: the expiration moves on
        context.Time.Advance(TimeSpan.FromSeconds(1));
        var second = await context.OpenCodiMD("/notes.md");
        await Assert.That(second.DocId).IsEqualTo(first.DocId);
        var wopiLock = JsonNode.Parse((await context.WopiLock("/notes.md"))!)!;
        await Assert.That(wopiLock["tocl"]!.AsObject().Count).IsEqualTo(2);

        using (var save = await context.Save(first, close: true))
        {
            await Assert.That(save.StatusCode).IsEqualTo(HttpStatusCode.Accepted);
        }

        context.Time.Advance(TimeSpan.FromSeconds(91));
        await context.Bridge.RunRound();
        await Assert.That(await context.WopiLock("/notes.md")).IsNotEqualTo("");
    }

    /// <summary>
    /// When the lock got lost (here expired), the save relocks the file first, as long as it did not change.
    /// </summary>
    [Test]
    public async Task CodiMD_RelocksLostLock()
    {
        var codimd = new CodiMDStub();
        await using var context = await Start(codimd.Handle);
        await File.WriteAllTextAsync(context.Host.PathOf("/notes.md"), "# Relock\n");
        var opened = await context.OpenCodiMD("/notes.md");
        context.Time.Advance(TimeSpan.FromMinutes(31));
        codimd.Notes[opened.DocId] = "# Relock\nEdited\n";
        using (var save = await context.Save(opened, close: false))
        {
            await Assert.That(save.StatusCode).IsEqualTo(HttpStatusCode.Accepted);
        }

        await context.Bridge.RunRound();
        await Assert.That(await File.ReadAllTextAsync(context.Host.PathOf("/notes.md"))).IsEqualTo("# Relock\nEdited\n");
        await Assert.That(await context.WopiLock("/notes.md")).IsNotEqualTo("");
    }

    /// <summary>
    /// When the lock got lost and the file changed meanwhile, the note is kept in the recovery space and the user told.
    /// </summary>
    [Test]
    public async Task CodiMD_RelockFailsAndRecovers()
    {
        var codimd = new CodiMDStub();
        await using var context = await Start(codimd.Handle);
        await File.WriteAllTextAsync(context.Host.PathOf("/notes.md"), "# Recover\n");
        var opened = await context.OpenCodiMD("/notes.md");
        context.Time.Advance(TimeSpan.FromMinutes(31));
        await File.WriteAllTextAsync(context.Host.PathOf("/notes.md"), "# Changed elsewhere\n");
        File.SetLastWriteTimeUtc(context.Host.PathOf("/notes.md"), context.Time.Now.UtcDateTime);
        codimd.Notes[opened.DocId] = "# Recover\nEdited\n";
        using (var save = await context.Save(opened, close: false))
        {
            await Assert.That(save.StatusCode).IsEqualTo(HttpStatusCode.Accepted);
        }

        await context.Bridge.RunRound();
        using (var again = await context.Save(opened, close: false))
        {
            await Assert.That(again.StatusCode).IsEqualTo(HttpStatusCode.InternalServerError);
            await Assert.That(await again.Content.ReadAsStringAsync()).Contains("The file was modified externally");
        }

        await Assert.That(await File.ReadAllTextAsync(context.Host.PathOf("/notes.md"))).IsEqualTo("# Changed elsewhere\n");
        await Assert.That(await File.ReadAllTextAsync(Directory.GetFiles(context.Host.RecoveryPath).Single())).IsEqualTo("# Recover\nEdited\n");
    }

    /// <summary>
    /// A document whose access token expired is dropped: by experience, these are left over browser sessions.
    /// </summary>
    [Test]
    public async Task CodiMD_ExpiredToken()
    {
        var codimd = new CodiMDStub();
        await using var context = await Start(codimd.Handle, "tokenvalidity = 60");
        await File.WriteAllTextAsync(context.Host.PathOf("/notes.md"), "# Expired\n");
        var opened = await context.OpenCodiMD("/notes.md");
        using (var save = await context.Save(opened, close: false))
        {
            await Assert.That(save.StatusCode).IsEqualTo(HttpStatusCode.Accepted);
        }

        context.Time.Advance(TimeSpan.FromSeconds(300));
        await context.Bridge.RunRound();
        var openFiles = await context.OpenFiles();
        await Assert.That((bool?) openFiles[opened.WopiSrc]!["toclose"]!["invalid-lock"]).IsTrue();
        context.Time.Advance(TimeSpan.FromSeconds(91));
        await context.Bridge.RunRound();
        await Assert.That((await context.OpenFiles()).Count).IsEqualTo(0);
    }

    /// <summary>
    /// A save request for a document the bridge does not know, as after a restart, repopulates its metadata.
    /// </summary>
    [Test]
    public async Task CodiMD_SaveUnknownDocument()
    {
        var codimd = new CodiMDStub();
        await using var context = await Start(codimd.Handle);
        await File.WriteAllTextAsync(context.Host.PathOf("/notes.md"), "# Restart\n");
        var opened = await context.OpenCodiMD("/notes.md");
        await using var restarted = await Start(codimd.Handle, "codimdurl = http://codimd");
        await File.WriteAllTextAsync(restarted.Host.PathOf("/notes.md"), "# Restart\n");
        // the save goes to a server that never saw the document open, so it has no lock either
        codimd.Notes[opened.DocId] = "# Restart\nEdited\n";
        using (var save = await restarted.Save(opened, close: false))
        {
            await Assert.That(save.StatusCode).IsEqualTo(HttpStatusCode.Accepted);
        }

        await Assert.That((await restarted.OpenFiles()).Count).IsEqualTo(1);
    }

    // Plugins

    /// <summary>
    /// The model supports one backend per app: a second app URL is refused.
    /// </summary>
    [Test]
    public async Task Plugin_SecondAppUrl()
    {
        var codimd = new CodiMDStub();
        await using var context = await Start(codimd.Handle);
        await File.WriteAllTextAsync(context.Host.PathOf("/notes.md"), "# Notes\n");
        await context.OpenCodiMD("/notes.md");
        using var response = await context.Open("/notes.md", appUrl: "http://other-codimd");
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.NotImplemented);
        await Assert.That(await response.Content.ReadAsStringAsync()).IsEqualTo("Bridged app CodiMD already configured with a different appurl");
    }

    /// <summary>
    /// An app that is not a CodiMD instance (no CodiMD metrics) fails to load.
    /// </summary>
    [Test]
    public async Task Plugin_NotCodiMD()
    {
        var codimd = new CodiMDStub
        {
            Available = false
        };
        await using var context = await Start(codimd.Handle);
        await File.WriteAllTextAsync(context.Host.PathOf("/notes.md"), "# Notes\n");
        using var response = await context.Open("/notes.md");
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.InternalServerError);
        await Assert.That(await response.Content.ReadAsStringAsync()).IsEqualTo("Failed to load WOPI bridge plugin for CodiMD");
    }

    [Test]
    public async Task Plugin_Unreachable()
    {
        await using var context = await Start(_ => throw new HttpRequestException("unreachable"));
        await File.WriteAllTextAsync(context.Host.PathOf("/notes.md"), "# Notes\n");
        using var response = await context.Open("/notes.md");
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.InternalServerError);
    }

    /// <summary>
    /// The apps can be configured, in which case their plugin is loaded at startup.
    /// </summary>
    [Test]
    public async Task Plugin_Configured()
    {
        var codimd = new CodiMDStub();
        await using var context = await Start(codimd.Handle, "codimdurl = http://codimd\ncodimdinturl = http://codimd");
        await File.WriteAllTextAsync(context.Host.PathOf("/notes.md"), "# Configured\n");
        var opened = await context.OpenCodiMD("/notes.md");
        await Assert.That(codimd.Notes[opened.DocId]).IsEqualTo("# Configured\n");
    }

    [Test]
    public async Task Plugin_ConfiguredButUnreachable() =>
        await Helpers.Throws<UnsupportedAppException>(() => TestServerHost.Start("codimdurl = http://codimd", apps: _ => throw new HttpRequestException("unreachable")));

    [Test]
    public async Task Save_MissingMetadata()
    {
        await using var context = await Start(new CodiMDStub().Handle);
        using var request = new HttpRequestMessage(HttpMethod.Post, "/wopi/bridge/doc?WOPISrc=x");
        request.Headers.Add("X-Efss-Bridged-App", "CodiMD");
        using var response = await context.Host.Client.SendAsync(request);
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.BadRequest);
        await Assert.That(await response.Content.ReadAsStringAsync()).StartsWith("""{"message": "Missing metadata, could not save.""");
    }

    [Test]
    public async Task List_Unauthorized()
    {
        await using var context = await Start(new CodiMDStub().Handle);
        using var response = await context.Host.Client.GetAsync("/wopi/bridge/list");
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.Unauthorized);
        using var request = new HttpRequestMessage(HttpMethod.Get, "/wopi/bridge/list");
        request.Headers.Add("Authorization", $"Bearer {TestServerHost.WopiSecret}");
        using var authorized = await context.Host.Client.SendAsync(request);
        await Assert.That(authorized.StatusCode).IsEqualTo(HttpStatusCode.OK);
    }

    // Etherpad

    /// <summary>
    /// A pad is imported into Etherpad on open, given the save webhook metadata, and saved back on close.
    /// </summary>
    [Test]
    public async Task Etherpad_OpenEditSave()
    {
        var etherpad = new EtherpadStub();
        await using var context = await Start(etherpad.Handle);
        await File.WriteAllTextAsync(context.Host.PathOf("/pad.epd"), """{"pad": "content"}""");
        using var response = await context.Open("/pad.epd", "Etherpad", "http://etherpad", apiKey: "etherpadkey");
        var appUrl = new Uri(await Context.AppUrl(response));
        var docId = appUrl.AbsolutePath["/p/".Length..];
        await Assert.That(appUrl.Query).IsEqualTo("?userName=Alice");
        await Assert.That(etherpad.Pads[docId]).IsEqualTo("""{"pad": "content"}""");
        await Assert.That(etherpad.Calls).Contains("createGroupIfNotExistsFor");
        var (wopiSrc, token) = etherpad.Metadata[docId];

        etherpad.Pads[docId] = """{"pad": "edited"}""";
        var opened = new Opened(appUrl.ToString(), docId, wopiSrc, token);
        using (var save = await context.Save(opened, close: true, app: "Etherpad"))
        {
            await Assert.That(save.StatusCode).IsEqualTo(HttpStatusCode.Accepted);
        }

        await context.Bridge.RunRound();
        await Assert.That(await File.ReadAllTextAsync(context.Host.PathOf("/pad.epd"))).IsEqualTo("""{"pad": "edited"}""");
    }

    /// <summary>
    /// A read-only open gives the read-only pad of a random pad id; an empty file gives an empty pad.
    /// </summary>
    [Test]
    public async Task Etherpad_ReadOnly()
    {
        var etherpad = new EtherpadStub();
        await using var context = await Start(etherpad.Handle);
        await File.WriteAllTextAsync(context.Host.PathOf("/pad.epd"), "");
        using var response = await context.Open("/pad.epd", "Etherpad", "http://etherpad", "VIEW_MODE_READ_ONLY", "etherpadkey");
        var appUrl = await Context.AppUrl(response);
        await Assert.That(appUrl).StartsWith("http://etherpad/p/r.");
        var pad = etherpad.Pads.Single();
        await Assert.That(pad.Key.Length).IsEqualTo(20);
        await Assert.That(pad.Value).IsEqualTo("{}");
    }

    /// <summary>
    /// An API error of Etherpad fails the plugin load.
    /// </summary>
    [Test]
    public async Task Etherpad_WrongApiKey()
    {
        var etherpad = new EtherpadStub();
        await using var context = await Start(etherpad.Handle);
        await File.WriteAllTextAsync(context.Host.PathOf("/pad.epd"), "{}");
        using var response = await context.Open("/pad.epd", "Etherpad", "http://etherpad", apiKey: "wrong");
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.InternalServerError);
        await Assert.That(await response.Content.ReadAsStringAsync()).IsEqualTo("Failed to load WOPI bridge plugin for Etherpad");
    }
}
