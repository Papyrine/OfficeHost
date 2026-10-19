/// <summary>
/// Real WOPI clients opening and saving documents through the WOPI endpoints. Each test starts its editor in Docker
/// (the first run pulls the image), launches the file with <see cref="WopiLauncher"/>, and drives the editor with
/// Playwright.
/// </summary>
public class EditorTests
{
    const string text = "Hello from the WOPI server";

    static Task<IBrowser> Browser(IPlaywright playwright) =>
        playwright.Chromium.LaunchAsync(
            new()
            {
                Headless = Environment.GetEnvironmentVariable("E2E_HEADED") == null
            });

    static async Task Screenshot(IPage page, string name)
    {
        var folder = Path.Combine(AppContext.BaseDirectory, "screenshots");
        Directory.CreateDirectory(folder);
        await page.ScreenshotAsync(
            new()
            {
                Path = Path.Combine(folder, $"{name}.png")
            });
    }

    /// <summary>
    /// Collabora loads the document (CheckFileInfo, GetFile, Lock), and saves the user's edit back (PutFile).
    /// </summary>
    [Test]
    [Timeout(900_000)]
    public async Task Collabora_OpenEditSave(Cancel cancel)
    {
        await using var collabora = await EditorContainer.Start(
            "collabora/code:26.04.4.2.1",
            9980,
            9980,
            new Dictionary<string, string>
            {
                // the WOPI hosts Collabora may call back, a regex matched against the host and port of the WOPISrc
                ["domain"] = @"host\.docker\.internal.*",
                // no TLS for the test, and no welcome dialog over the document
                ["extra_params"] = "--o:ssl.enable=false --o:ssl.termination=false --o:welcome.enable=false"
            },
            TimeSpan.FromMinutes(3),
            cancel);
        await using var server = await RunningServer.Start(_ => _.DiscoveryUrls.Add(collabora.DiscoveryUrl));
        var path = Path.Combine(server.StoragePath, "test.docx");
        await File.WriteAllBytesAsync(path, Docx.Create(text), cancel);
        var (fileId, launch) = await server.Launch("test.docx", collabora.DiscoveryUrl);

        using var playwright = await Playwright.CreateAsync();
        await using var browser = await Browser(playwright);
        var page = await browser.NewPageAsync();
        try
        {
            await EditorPage.Open(page, launch);
            await server.WaitForRequest($"POST /wopi/files/{fileId} 200", TimeSpan.FromMinutes(2), cancel);
            await Assert.That(server.Requests).Contains($"GET /wopi/files/{fileId}/contents 200");
            await Assert.That(await server.Service<IWopiLockStore>().Get(fileId, cancel)).IsNotNull();

            // type at the start of the document, and save
            var canvas = page.Locator("#document-canvas");
            await canvas.WaitForAsync(
                new()
                {
                    Timeout = 120_000
                });
            await page.WaitForTimeoutAsync(2000);
            // a dialog over the document, as Collabora's welcome on first use, would take the click
            var welcome = page.Locator(".iframe-welcome-modal");
            if (await welcome.IsVisibleAsync())
            {
                await page.Keyboard.PressAsync("Escape");
                await welcome.WaitForAsync(
                    new()
                    {
                        State = WaitForSelectorState.Hidden,
                        Timeout = 10_000
                    });
            }

            await canvas.ClickAsync(
                new()
                {
                    Position = new()
                    {
                        X = 300,
                        Y = 150
                    }
                });
            await page.Keyboard.PressAsync("Control+Home");
            await page.Keyboard.TypeAsync("Edited in Collabora. ");
            await page.Keyboard.PressAsync("Control+s");

            var deadline = DateTime.UtcNow + TimeSpan.FromMinutes(2);
            string? saved;
            while ((saved = Docx.Text(path)) == null ||
                   !saved.Contains("Edited in Collabora."))
            {
                if (DateTime.UtcNow > deadline)
                {
                    throw new TimeoutException($"The edit was not saved: {saved}");
                }

                await Task.Delay(1000, cancel);
            }

            await Assert.That(saved).Contains(text);
            await Assert.That(server.Requests).Contains($"POST /wopi/files/{fileId}/contents 200");
        }
        catch
        {
            await Screenshot(page, nameof(Collabora_OpenEditSave));
            await Console.Error.WriteLineAsync(string.Join('\n', server.Requests));
            await Console.Error.WriteLineAsync(server.Log.ToString());
            await Console.Error.WriteLineAsync(await collabora.Logs(cancel));
            throw;
        }
    }

    /// <summary>
    /// ONLYOFFICE loads the document, with every request it makes validated against the proof keys of its discovery:
    /// RequireProof refuses any unsigned or mismatching request. Saving is not tested, as ONLYOFFICE saves on its own
    /// schedule.
    /// </summary>
    [Test]
    [Timeout(1_500_000)]
    public async Task OnlyOffice_OpenWithProofKeys(Cancel cancel)
    {
        await using var onlyOffice = await EditorContainer.Start(
            "onlyoffice/documentserver:9.4.0.1",
            9981,
            80,
            new Dictionary<string, string>
            {
                ["WOPI_ENABLED"] = "true",
                // the JWT of ONLYOFFICE's own document API, unrelated to WOPI
                ["JWT_ENABLED"] = "false"
            },
            TimeSpan.FromMinutes(8),
            cancel);
        await using var server = await RunningServer.Start(
            _ =>
            {
                _.DiscoveryUrls.Add(onlyOffice.DiscoveryUrl);
                _.RequireProof = true;
            });
        await File.WriteAllBytesAsync(Path.Combine(server.StoragePath, "test.docx"), Docx.Create(text), cancel);
        var (fileId, launch) = await server.Launch("test.docx", onlyOffice.DiscoveryUrl);

        using var playwright = await Playwright.CreateAsync();
        await using var browser = await Browser(playwright);
        var page = await browser.NewPageAsync();
        try
        {
            await EditorPage.Open(page, launch);
            // GetFile: the document went to the editor
            await server.WaitForRequest($"GET /wopi/files/{fileId}/contents 200", TimeSpan.FromMinutes(3), cancel);
            await Assert.That(server.Requests.Any(_ => _.EndsWith(" 500"))).IsFalse();
            await Assert.That(server.Log.ToString()).DoesNotContain("invalid proof");
        }
        catch
        {
            await Screenshot(page, nameof(OnlyOffice_OpenWithProofKeys));
            await Console.Error.WriteLineAsync(string.Join('\n', server.Requests));
            await Console.Error.WriteLineAsync(server.Log.ToString());
            await Console.Error.WriteLineAsync(await onlyOffice.Logs(cancel));
            throw;
        }
    }
}
