/// <summary>
/// What a replay found: a line per exchange or sub-response, and the sub-responses that differ by their label.
/// </summary>
/// <param name="Contents">The Cells after which the files of both copies of the content were compared.</param>
record ReplayResult(List<string> Lines, int Same, Dictionary<string, int> Different, int Skipped, int Contents)
{
    public override string ToString() => string.Join(Environment.NewLine, Lines);
}

/// <summary>
/// Replays a recording to an OfficeHost on a test server with <see cref="ShadowReplay"/>, the engine of the shadow
/// (docs/replay.md), and reports it line by line. The documents start as a file of the same name in the files folder,
/// or as the first full content query of the recording answered them; one made in the recording, by a Save As, does
/// not.
/// </summary>
class Replayer : IShadowTarget, IShadowFiles, IAsyncDisposable
{
    TestServerHost host = null!;
    ManualTime time = new();
    Dictionary<(string Path, string User), ShadowFile> files = [];
    List<string> lines = [];
    int same;
    Dictionary<string, int> different = [];
    int skipped;
    int contents;

    // where the files of a content query are written when they differ
    string? output;

    /// <param name="files">The documents to start from, by name.</param>
    /// <param name="output">Where to write the two files of a content query whose files differ, in <c>differences</c>.</param>
    public static async Task<ReplayResult> Run(List<Exchange> exchanges, string? files = null, string? output = null)
    {
        await using var replayer = new Replayer
        {
            output = output
        };
        replayer.host = await TestServerHost.Start(time: replayer.time);
        await replayer.Seed(exchanges, files);
        var replay = new ShadowReplay(replayer, replayer, ShadowRule.Defaults);
        foreach (var exchange in exchanges)
        {
            var replayed = await replay.Replay(exchange);
            if (replayed.NotReplayed is { } reason)
            {
                if (ShadowComparer.Kind(exchange.Request) != null)
                {
                    replayer.lines.Add($"#{exchange.Number} {ShadowComparer.Kind(exchange.Request)}: not replayed, {reason}");
                }

                replayer.skipped++;
                continue;
            }

            foreach (var finding in replayed.Findings)
            {
                await replayer.Report(finding);
            }
        }

        var differences = replayer.different.Values.Sum();
        replayer.lines.Add($"{replayer.same} the same, {differences} different, {replayer.skipped} requests not replayed, the content compared after {replayer.contents} Cells");
        foreach (var (label, count) in replayer.different.OrderBy(_ => _.Key))
        {
            replayer.lines.Add($"  {label}: {count} different");
        }

        return new(replayer.lines, replayer.same, replayer.different, replayer.skipped, replayer.contents);
    }

    /// <summary>
    /// Writes each document of the cell storage requests to the storage, at its path in its host: the file of its name
    /// in the files folder, or the file of the first content query SharePoint answered in full. A document whose first
    /// request puts it, made by a Save As in the recording, is not written.
    /// </summary>
    async Task Seed(List<Exchange> exchanges, string? files)
    {
        var documents = new Dictionary<string, (ShadowDocument Document, bool Made)>();
        foreach (var exchange in exchanges.Where(_ => ShadowComparer.Kind(_.Request) == "cellstorage"))
        {
            if (SoapBody.TryRead(exchange.Request) is not { } request)
            {
                continue;
            }

            foreach (var (document, makes) in ShadowComparer.Documents(request))
            {
                documents.TryAdd(document.Key, (document, makes));
            }
        }

        foreach (var (document, made) in documents.Values)
        {
            byte[]? content = null;
            if (files != null &&
                File.Exists(Path.Combine(files, document.Name)))
            {
                content = await File.ReadAllBytesAsync(Path.Combine(files, document.Name));
            }

            if (content == null &&
                made)
            {
                lines.Add($"Not seeded {document.Name}: made in the recording");
                continue;
            }

            content ??= exchanges
                .Where(_ => ShadowComparer.Kind(_.Request) == "cellstorage" && _.Response.Status == 200)
                .SelectMany(_ => ShadowComparer.Contents(_, document.Key))
                .FirstOrDefault();
            if (content == null)
            {
                lines.Add($"No content for {document.Name}: not seeded (put the file in the files folder to seed it)");
                continue;
            }

            var path = host.PathOf(document.Path);
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            await File.WriteAllBytesAsync(path, content);
            lines.Add($"Seeded {document.Name}, {content.Length} bytes");
        }
    }

    /// <summary>
    /// The user of OfficeHost of a user of SharePoint: the name without its domain or claim prefix
    /// (<c>i:0#.w|lab\alice</c> is <c>alice</c>).
    /// </summary>
    static string UserOf(string recorded)
    {
        var name = recorded.Split('|')[^1].Split('\\')[^1].ToLowerInvariant();
        if (name.Length == 0)
        {
            return "alice";
        }

        return name;
    }

    public async Task<ShadowFile?> Open(ShadowDocument document, string user, Cancel cancel)
    {
        var userId = UserOf(user);
        if (files.TryGetValue((document.Path, userId), out var file))
        {
            return file;
        }

        if (!File.Exists(host.PathOf(document.Path)))
        {
            return null;
        }

        var (fileId, token) = await host.Open(document.Path, userId: userId);
        file = new(fileId, $"/fsshttp/{token}");
        files[(document.Path, userId)] = file;
        return file;
    }

    public async Task<HttpMessage> Send(string path, HttpMessage recorded, byte[] body, DateTimeOffset? at, Cancel cancel)
    {
        if (at != null)
        {
            time.Now = at.Value;
        }

        using var request = new HttpRequestMessage(HttpMethod.Post, path);
        request.Content = new ByteArrayContent(body);
        request.Content.Headers.TryAddWithoutValidation("Content-Type", recorded.ContentType);
        if (recorded.Header("SOAPAction") is { } action)
        {
            request.Headers.TryAddWithoutValidation("SOAPAction", action);
        }

        using var response = await host.Client.SendAsync(request, cancel);
        var headers = response.Content.Headers.Select(_ => (_.Key, string.Join(", ", _.Value))).ToList();
        return new($"HTTP/1.1 {(int) response.StatusCode} {response.ReasonPhrase}", headers, await response.Content.ReadAsByteArrayAsync(cancel));
    }

    async Task Report(ShadowFinding finding)
    {
        var number = finding.Exchange.Number;
        if (finding.ContentCompared)
        {
            contents++;
        }

        if (finding.Outcome == ShadowOutcome.Same)
        {
            same++;
            lines.Add($"#{number} {finding.Label}: same {finding.Detail}".TrimEnd());
            return;
        }

        different[finding.Label] = different.GetValueOrDefault(finding.Label) + 1;
        var outcome = finding.Outcome switch
        {
            ShadowOutcome.Different => "DIFFERENT",
            ShadowOutcome.KnownDifference => "KNOWN DIFFERENCE",
            ShadowOutcome.NotCompared => "NOT COMPARED",
            ShadowOutcome.OfficeHostFailed => "OFFICEHOST FAILED",
            _ => finding.Outcome.ToString().ToUpperInvariant()
        };
        var reason = "";
        if (finding.Reason != null)
        {
            reason = $", {finding.Reason}";
        }

        lines.Add($"#{number} {finding.Label}: {outcome} (SharePoint / OfficeHost{reason})");
        foreach (var difference in finding.Differences)
        {
            lines.Add($"    {difference}");
        }

        if (output != null &&
            finding.Files is { } files)
        {
            var folder = Directory.CreateDirectory(Path.Combine(output, "differences")).FullName;
            var extension = Path.GetExtension(finding.Document ?? "");
            var prefix = Path.Combine(folder, $"{number:0000}-{different[finding.Label]}");
            await File.WriteAllBytesAsync($"{prefix}-sharepoint{extension}", files.SharePoint);
            await File.WriteAllBytesAsync($"{prefix}-officehost{extension}", files.OfficeHost);
            lines.Add($"    files written to {prefix}-*{extension}");
        }
    }

    public async ValueTask DisposeAsync()
    {
        if (host != null)
        {
            await host.DisposeAsync();
        }
    }
}
