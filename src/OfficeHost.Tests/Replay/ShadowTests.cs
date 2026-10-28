/// <summary>
/// The shadow (docs/shadow.md) end to end, in process: an OfficeHost recorded as SharePoint would be by the module, its
/// exchanges posted to the ingest of a shadow application as the module posts them, and the shadow's log.
/// </summary>
public class ShadowTests
{
    internal const string Key = "the key of the shadow of the tests";

    /// <summary>
    /// The content of the documents, as SharePoint would give it: the file as the recording started.
    /// </summary>
    sealed class Content(byte[] file) : IShadowContent
    {
        public Task<ShadowContent?> Get(ShadowDocument document, Cancel cancel = default) =>
            Task.FromResult<ShadowContent?>(new(file, null));
    }

    /// <summary>
    /// A session of two Word clients with an OfficeHost, recorded as SharePoint's would be: the open, a save, the other
    /// client's query after it, the exits. Its clock, the recorded times, is <paramref name="time"/>.
    /// </summary>
    internal static async Task<(List<Exchange> Exchanges, byte[] File)> Record(TempDirectory folder, ManualTime? time = null)
    {
        var original = Zips.Create(
            ("word/document.xml", Zips.Random(3000, 1)),
            ("word/styles.xml", Zips.Random(3000, 2)));
        await using var recorded = await TestServerHost.Start(time: time);
        await File.WriteAllBytesAsync(recorded.PathOf("doc.docx"), original);
        var (_, token) = await recorded.Open("doc.docx");
        recorded.Record(folder);
        var word = new WordClient(recorded, token);
        var other = new WordClient(recorded, token);
        var join = (await word.Send(word.Join())).Single();
        var transition = join.Attribute("TransitionID")!;
        await word.Sync();
        await word.IsOnlyClient(transition);
        await other.Send(other.Join());
        await other.Sync();
        await word.Coauth("GetCoauthoringStatus");
        await word.Save(Zips.Create(("word/document.xml", Zips.Random(3000, 3)), ("word/styles.xml", Zips.Random(3000, 2))));
        await other.Sync();
        await other.Coauth("ExitCoauthoring");
        await word.Coauth("ExitCoauthoring");
        return (Exchange.ReadFolder(folder), original);
    }

    internal static Task<TestServerHost> StartShadow(byte[] file, Action<OfficeHostShadowOptions>? configure = null, Func<IOfficeStorage, IOfficeStorage>? wrapStorage = null) =>
        TestServerHost.Start(
            wrapStorage: wrapStorage,
            services: services =>
            {
                services.AddOfficeHostShadow(options =>
                {
                    options.ApiKey = Key;
                    options.MaximumDelay = TimeSpan.FromMinutes(1);
                    options.Levels[ShadowOutcome.Same] = LogLevel.Information;
                    configure?.Invoke(options);
                });
                services.AddSingleton<IShadowDocuments>(_ => new FileSystemShadowDocuments(_.GetRequiredService<FileSystemStorage>()));
                services.AddSingleton<IShadowContent>(new Content(file));
            },
            endpoints: app => app.MapOfficeHostShadow());

    /// <summary>
    /// Posts exchanges and tombstones to the ingest as an instance of the module does (<see cref="ShadowProtocol"/>).
    /// </summary>
    static async Task<HttpResponseMessage> Post(
        TestServerHost shadow,
        IEnumerable<Exchange> exchanges,
        Guid instance,
        DateTimeOffset? watermark = null,
        string sentKey = Key,
        IEnumerable<(ShadowDocument Document, long Sequence, DateTimeOffset Started)>? tombstones = null)
    {
        var content = new MultipartContent("mixed");
        var latest = DateTimeOffset.MinValue;
        foreach (var exchange in exchanges)
        {
            var started = exchange.Time!.Value;
            if (started > latest)
            {
                latest = started;
            }

            var request = exchange.Request.ToArray();
            var part = new ByteArrayContent([.. request, .. exchange.Response.ToArray()]);
            part.Headers.TryAddWithoutValidation(ShadowProtocol.Sequence, exchange.Number.ToString());
            part.Headers.TryAddWithoutValidation(ShadowProtocol.Started, started.ToString("O"));
            part.Headers.TryAddWithoutValidation(ShadowProtocol.Kind, "exchange");
            part.Headers.TryAddWithoutValidation(ShadowProtocol.RequestLength, request.Length.ToString());
            content.Add(part);
        }

        foreach (var (document, sequence, started) in tombstones ?? [])
        {
            var part = new ByteArrayContent([]);
            part.Headers.TryAddWithoutValidation(ShadowProtocol.Sequence, sequence.ToString());
            part.Headers.TryAddWithoutValidation(ShadowProtocol.Started, started.ToString("O"));
            part.Headers.TryAddWithoutValidation(ShadowProtocol.Kind, "tombstone");
            part.Headers.TryAddWithoutValidation(ShadowProtocol.Document, document.Url.ToString());
            part.Headers.TryAddWithoutValidation(ShadowProtocol.Reason, "too large");
            content.Add(part);
        }

        using var message = new HttpRequestMessage(HttpMethod.Post, "/officehost-shadow/exchanges");
        // a heartbeat has no body
        if (content.Any())
        {
            message.Content = content;
        }

        message.Headers.Add(ShadowProtocol.Key, sentKey);
        message.Headers.Add(ShadowProtocol.Instance, instance.ToString());
        message.Headers.Add(ShadowProtocol.Server, "sharepoint");
        message.Headers.Add(ShadowProtocol.Watermark, (watermark ?? latest.AddSeconds(1)).ToString("O"));
        return await shadow.Client.SendAsync(message);
    }

    /// <summary>
    /// The findings the shadow logged (<c>Level ShadowProcessor: Outcome: …</c>), once there are <paramref name="count"/>.
    /// </summary>
    internal static async Task<List<string>> Findings(TestServerHost shadow, int count)
    {
        for (var attempt = 0; attempt < 300; attempt++)
        {
            var findings = shadow.Log.Lines
                .Where(_ => _.Contains(" ShadowProcessor: ", StringComparison.Ordinal) && !_.Contains("Shadow since", StringComparison.Ordinal))
                .ToList();
            if (findings.Count >= count)
            {
                return findings;
            }

            await Task.Delay(50);
        }

        throw new TimeoutException($"Not {count} findings:\n{shadow.Log}");
    }

    /// <summary>
    /// The recorded session as one batch: the document started from SharePoint's content at the join (no content query
    /// in it), then every sub-response the same, the save made again for OfficeHost (the cells translated), at the
    /// levels of the options.
    /// </summary>
    [Test]
    public async Task Same()
    {
        using var folder = new TempDirectory();
        var (exchanges, file) = await Record(folder);
        await using var shadow = await StartShadow(file);

        using var response = await Post(shadow, exchanges, Guid.NewGuid());
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.Accepted);
        var findings = await Findings(shadow, 10);
        await Task.Delay(200);
        Console.WriteLine(shadow.Log);

        await Assert.That(shadow.Log.Lines.Count(_ => _.StartsWith("Information ShadowProcessor: Same: ", StringComparison.Ordinal))).IsEqualTo(10);
        await Assert.That(findings).Contains(_ => _.Contains("Same: Cell PutChanges (translated) of localhost/fsshttp/doc.docx"));
        await Assert.That(shadow.Log.Lines).DoesNotContain(_ => _.StartsWith("Warning", StringComparison.Ordinal) || _.StartsWith("Error", StringComparison.Ordinal));
    }

    /// <summary>
    /// An answer of SharePoint that OfficeHost does not give is Different, at Warning; a rule of the options that
    /// expects it makes it a KnownDifference, at its level.
    /// </summary>
    [Test]
    public async Task Different_And_KnownDifference()
    {
        using var folder = new TempDirectory();
        var (_, file) = await Record(folder);
        var status = Path.Combine(folder, "0006.response.http");
        await File.WriteAllTextAsync(status, (await File.ReadAllTextAsync(status)).Replace("CoauthStatus=\"Coauthoring\"", "CoauthStatus=\"Alone\""));
        var exchanges = Exchange.ReadFolder(folder);

        await using (var shadow = await StartShadow(file))
        {
            await Post(shadow, exchanges, Guid.NewGuid());
            var findings = await Findings(shadow, 10);
            await Assert.That(findings).Contains(_ => _.StartsWith("Warning ShadowProcessor: Different: Coauth GetCoauthoringStatus of localhost/fsshttp/doc.docx", StringComparison.Ordinal) &&
                                                      _.Contains("@CoauthStatus: Alone / Coauthoring"));
        }

        await using (var shadow = await StartShadow(file, options =>
                     {
                         options.KnownDifferences.Add(new("^Coauth GetCoauthoringStatus$", "^@CoauthStatus$", "the test expects it"));
                         options.Levels[ShadowOutcome.KnownDifference] = LogLevel.Information;
                     }))
        {
            await Post(shadow, exchanges, Guid.NewGuid());
            var findings = await Findings(shadow, 10);
            await Assert.That(findings).Contains(_ => _.StartsWith("Information ShadowProcessor: KnownDifference: Coauth GetCoauthoringStatus", StringComparison.Ordinal) &&
                                                      _.Contains("(the test expects it)"));
            await Assert.That(findings).DoesNotContain(_ => _.Contains("Different:"));
        }
    }

    /// <summary>
    /// OfficeHost failing to write the save, where SharePoint applied it: OfficeHostFailed, at Error.
    /// </summary>
    [Test]
    public async Task OfficeHostFailed()
    {
        using var folder = new TempDirectory();
        var (exchanges, file) = await Record(folder);
        FaultyStorage? storage = null;
        await using var shadow = await StartShadow(file, wrapStorage: _ => storage = new(_));
        storage!.Fault = _ => _ == nameof(IOfficeStorage.Write) ? new IOException("the disk is full") : null;

        await Post(shadow, exchanges, Guid.NewGuid());
        var findings = await Findings(shadow, 10);
        await Task.Delay(200);
        Console.WriteLine(shadow.Log);

        await Assert.That(findings).Contains(_ => _.StartsWith("Error ShadowProcessor: OfficeHostFailed: ", StringComparison.Ordinal));
    }

    /// <summary>
    /// A tombstone of the module (an exchange too large to send) makes its document untrusted: a difference after it is
    /// not compared, and the tombstone is logged as TapDropped.
    /// </summary>
    [Test]
    public async Task Tombstone()
    {
        using var folder = new TempDirectory();
        var (_, file) = await Record(folder);
        var status = Path.Combine(folder, "0006.response.http");
        await File.WriteAllTextAsync(status, (await File.ReadAllTextAsync(status)).Replace("CoauthStatus=\"Coauthoring\"", "CoauthStatus=\"Alone\""));
        var exchanges = Exchange.ReadFolder(folder);
        await using var shadow = await StartShadow(file, options => options.Levels[ShadowOutcome.NotCompared] = LogLevel.Information);

        var document = ShadowDocument.From("https://localhost/fsshttp/doc.docx")!;
        var started = exchanges[3].Time!.Value.AddTicks(1);
        await Post(shadow, exchanges, Guid.NewGuid(), tombstones: [(document, 100, started)]);
        var findings = await Findings(shadow, 11);

        await Assert.That(findings).Contains(_ => _.StartsWith("Information ShadowProcessor: TapDropped: tombstone of localhost/fsshttp/doc.docx", StringComparison.Ordinal));
        // the difference after it is not compared: the document's state is no longer known
        await Assert.That(findings).Contains(_ => _.StartsWith("Information ShadowProcessor: NotCompared: Coauth GetCoauthoringStatus", StringComparison.Ordinal) &&
                                                  _.Contains("an exchange missing: too large"));
        await Assert.That(findings).DoesNotContain(_ => _.Contains("Different:"));
    }

    /// <summary>
    /// Exchanges of two servers, the later one first: each server's watermark holds the later one back until the
    /// earlier one has come, and they are replayed in the order SharePoint started them.
    /// </summary>
    [Test]
    public async Task Order_ByWatermarks()
    {
        using var folder = new TempDirectory();
        var (exchanges, file) = await Record(folder);
        await using var shadow = await StartShadow(file);
        var first = Guid.NewGuid();
        var second = Guid.NewGuid();
        var before = exchanges[0].Time!.Value.AddSeconds(-1);

        // the first server is active, with nothing yet: its watermark holds back what the second sends
        await Post(shadow, [], first, before);
        await Post(shadow, exchanges.Skip(1), second);
        await Task.Delay(1500);
        await Assert.That(shadow.Log.Lines).DoesNotContain(_ => _.Contains(" ShadowProcessor: ", StringComparison.Ordinal));

        await Post(shadow, exchanges.Take(1), first, exchanges[^1].Time!.Value.AddSeconds(1));
        var findings = await Findings(shadow, 10);
        await Assert.That(findings[0]).Contains("Same: Coauth JoinCoauthoring of localhost/fsshttp/doc.docx");
        await Assert.That(findings.Count(_ => _.Contains("Same: "))).IsEqualTo(10);
    }

    /// <summary>
    /// OfficeHost replays each exchange at the time SharePoint received it: a coauthoring session that expired between
    /// two exchanges on SharePoint expires on OfficeHost too, however soon after each other the shadow replays them.
    /// </summary>
    [Test]
    public async Task ReplayClock()
    {
        using var folder = new TempDirectory();
        var time = new ManualTime();
        var original = Zips.Create(("word/document.xml", Zips.Random(3000, 1)));
        await using (var recorded = await TestServerHost.Start(time: time))
        {
            await File.WriteAllBytesAsync(recorded.PathOf("doc.docx"), original);
            var (_, token) = await recorded.Open("doc.docx");
            recorded.Record(folder);
            var word = new WordClient(recorded, token);
            await word.Send(word.Join());
            // past the hour of the session
            time.Advance(TimeSpan.FromHours(2));
            var status = await word.Coauth("GetCoauthoringStatus");
            await Assert.That(status.Attribute("CoauthStatus")).IsEqualTo("None");
        }

        await using var shadow = await StartShadow(original);
        await Post(shadow, Exchange.ReadFolder(folder), Guid.NewGuid());
        var findings = await Findings(shadow, 2);

        await Assert.That(findings).Contains(_ => _.Contains("Same: Coauth GetCoauthoringStatus"));
    }

    [Test]
    public async Task WrongKey()
    {
        using var folder = new TempDirectory();
        var (exchanges, file) = await Record(folder);
        await using var shadow = await StartShadow(file);

        using var response = await Post(shadow, exchanges, Guid.NewGuid(), sentKey: "another key");
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.Unauthorized);
    }

    /// <summary>
    /// A batch beyond the capacity of the inbox: 503 with Retry-After, so that the module backs off.
    /// </summary>
    [Test]
    public async Task Full()
    {
        using var folder = new TempDirectory();
        var (exchanges, file) = await Record(folder);
        await using var shadow = await StartShadow(file, options => options.QueueCapacity = 2);

        using var response = await Post(shadow, exchanges, Guid.NewGuid());
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.ServiceUnavailable);
        await Assert.That(response.Headers.RetryAfter).IsNotNull();
    }

    /// <summary>
    /// The sampling of the options goes to the module with every answer.
    /// </summary>
    [Test]
    public async Task Sample()
    {
        await using var shadow = await StartShadow([], options => options.SamplePercent = 25);
        using var response = await Post(shadow, [], Guid.NewGuid(), DateTimeOffset.UtcNow);
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.Accepted);
        await Assert.That(string.Join(",", response.Headers.GetValues(ShadowProtocol.Sample))).IsEqualTo("25");
    }
}
