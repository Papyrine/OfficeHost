/// <summary>
/// The module in SharePoint (OfficeHost.SharePoint) but for what needs System.Web, its sources linked here: sampling,
/// the bounded copies and queue, the watermark, the sender, the folder sink, and its batches as the shadow reads them.
/// </summary>
public class TapTests
{
    const string envelope = """<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/"><s:Body><RequestVersion Version="2" MinorVersion="2" xmlns="http://schemas.microsoft.com/sharepoint/soap/" /><RequestCollection CorrelationId="{6F7E7A3C-0000-4000-8000-000000000001}" xmlns="http://schemas.microsoft.com/sharepoint/soap/"><Request Url="http://spse.mshome.net/Shared%20Documents/Budget%20Q1.xlsx?x=1&amp;y=2" RequestToken="1"><SubRequest Type="WhoAmI" SubRequestToken="1" /></Request></RequestCollection></s:Body></s:Envelope>""";

    /// <summary>
    /// The document of a request from the start of its body, its key the shadow's, and a sampling the same for a document
    /// every time, about the percentage of documents.
    /// </summary>
    [Test]
    public async Task Sampling()
    {
        var document = TapSampler.Document($"\r\n--boundary\r\nContent-ID: <http://tempuri.org/xml/0>\r\n\r\n{envelope}")!;
        await Assert.That(document).IsEqualTo("http://spse.mshome.net/Shared%20Documents/Budget%20Q1.xlsx?x=1&y=2");
        await Assert.That(TapSampler.Key(document)).IsEqualTo(ShadowDocument.From(document)!.Key);
        await Assert.That(TapSampler.Document("<html>not a request</html>")).IsNull();

        var keys = Enumerable.Range(0, 10_000).Select(_ => $"spse/shared documents/{_}.docx").ToList();
        var sampled = keys.Count(_ => TapSampler.Sampled(_, 25));
        await Assert.That(sampled).IsBetween(2_300, 2_700);
        await Assert.That(keys.All(_ => TapSampler.Sampled(_, 25) == TapSampler.Sampled(_, 25))).IsTrue();
        await Assert.That(keys.All(_ => TapSampler.Sampled(_, 100))).IsTrue();
        await Assert.That(keys.Any(_ => TapSampler.Sampled(_, 0))).IsFalse();
    }

    /// <summary>
    /// A copy stops at its limit, or when the budget of all copies is spent, and gives its chunks back.
    /// </summary>
    [Test]
    public async Task Buffer_Bounded()
    {
        var budget = new TapBudget(256 * 1024);
        var bytes = Enumerable.Range(0, 100_000).Select(_ => (byte) _).ToArray();
        var copy = new TapBuffer(budget, 1024 * 1024);
        copy.Write(bytes, 0, bytes.Length);
        await Assert.That(copy.ToArray().SequenceEqual(bytes)).IsTrue();
        await Assert.That(budget.Used).IsEqualTo(2 * 64 * 1024);

        var other = new TapBuffer(budget, 1024 * 1024);
        other.Write(bytes, 0, bytes.Length);
        other.Write(bytes, 0, bytes.Length);
        await Assert.That(other.Overflowed).IsTrue();
        await Assert.That(budget.Used).IsEqualTo(2 * 64 * 1024);

        var limited = new TapBuffer(budget, 50_000);
        limited.Write(bytes, 0, bytes.Length);
        await Assert.That(limited.Overflowed).IsTrue();

        copy.Release();
        await Assert.That(budget.Used).IsEqualTo(0);
    }

    /// <summary>
    /// The queue never takes more than its capacity, counts the tombstones beyond theirs, and gives as watermark the
    /// earliest start of what waits or is in flight.
    /// </summary>
    [Test]
    public async Task Queue_BoundedWithWatermark()
    {
        var queue = new TapQueue(capacity: 2, tombstoneCapacity: 1);
        var start = new DateTime(2026, 10, 28, 6, 0, 0, DateTimeKind.Utc);
        TapItem Item(int second) => TapItem.Exchange(second, start.AddSeconds(second), null, [1], null, [2], null);

        await Assert.That(queue.TryAdd(Item(3))).IsTrue();
        await Assert.That(queue.TryAdd(Item(4))).IsTrue();
        await Assert.That(queue.TryAdd(Item(5))).IsFalse();
        queue.AddTombstone(TapItem.Tombstone(6, start.AddSeconds(6), "doc", "too large"));
        queue.AddTombstone(TapItem.Tombstone(7, start.AddSeconds(7), "doc", "too large"));
        queue.Started(8, start.AddSeconds(1));

        var batch = queue.Take(1, long.MaxValue, out var watermark, out var dropped);
        await Assert.That(batch.Single().Sequence).IsEqualTo(3);
        // the request in flight started first
        await Assert.That(watermark).IsEqualTo(start.AddSeconds(1));
        await Assert.That(dropped).IsEqualTo(1);

        queue.Ended(8);
        batch = queue.Take(10, long.MaxValue, out watermark, out _);
        await Assert.That(batch.Select(_ => _.Sequence)).IsEquivalentTo([4L, 6L]);
        await Assert.That(watermark).IsGreaterThan(start.AddSeconds(6));

        queue.Return(batch, 0);
        await Assert.That(queue.Count).IsEqualTo(2);
    }

    sealed class Sink(Func<int, TapSent> answer) : ITapSink
    {
        public int Calls { get; private set; }
        public List<List<TapItem>> Batches { get; } = [];

        public TapSent Send(IReadOnlyList<TapItem> items, DateTime watermark, long dropped, bool final)
        {
            Calls++;
            Batches.Add(items.ToList());
            return answer(Calls);
        }
    }

    /// <summary>
    /// The sender: a sink that fails three times opens it, and what waits becomes tombstones; the sink answering again
    /// closes it; 503 keeps the exchanges; the shadow's sampling is taken.
    /// </summary>
    [Test]
    public async Task Sender()
    {
        var queue = new TapQueue();
        var start = DateTime.UtcNow;
        var budget = new TapBudget(1024 * 1024);
        var body = new TapBuffer(budget, 1024 * 1024);
        body.Write([1, 2, 3], 0, 3);
        queue.TryAdd(TapItem.Exchange(1, start, "http://sp/doc.docx", [1], body, [2], null));
        var failing = new Sink(_ => new() { Error = "unreachable" });
        var sender = new TapSender(queue, failing, 100, _ => { });

        for (var attempt = 0; attempt < 3; attempt++)
        {
            sender.SendAll(final: true);
        }

        await Assert.That(sender.Open).IsTrue();
        await Assert.That(budget.Used).IsEqualTo(0);
        var buried = queue.Take(10, long.MaxValue, out _, out _).Single();
        await Assert.That(buried.IsTombstone).IsTrue();
        await Assert.That(buried.Document).IsEqualTo("http://sp/doc.docx");

        queue.Return([buried], 0);
        var answering = new Sink(_ => new() { Sent = true, SamplePercent = 10 });
        sender = new(queue, answering, 100, _ => { });
        sender.SendAll(final: true);
        await Assert.That(sender.Open).IsFalse();
        await Assert.That(sender.SamplePercent).IsEqualTo(10);
        await Assert.That(queue.Count).IsEqualTo(0);

        queue.TryAdd(TapItem.Exchange(2, start, null, [1], null, [2], null));
        var full = new Sink(_ => new() { RetryAfter = TimeSpan.FromSeconds(30) });
        sender = new(queue, full, 100, _ => { });
        sender.SendAll(final: true);
        await Assert.That(queue.Count).IsEqualTo(1);
        await Assert.That(queue.Take(1, long.MaxValue, out _, out _).Single().IsTombstone).IsFalse();
    }

    static List<TapItem> Items(List<Exchange> exchanges, TapBudget budget) =>
        exchanges
            .Select(_ =>
            {
                // the head as the module writes it, the body copied
                var request = _.Request.ToArray();
                var head = request.Length - _.Request.Body.Length;
                var body = new TapBuffer(budget, long.MaxValue);
                body.Write(_.Request.Body, 0, _.Request.Body.Length);
                return TapItem.Exchange(_.Number, _.Time!.Value.UtcDateTime, null, request[..head], body, _.Response.ToArray(), null);
            })
            .ToList();

    /// <summary>
    /// The folder sink records as the replay reads: the same exchanges.
    /// </summary>
    [Test]
    public async Task FolderSink_Replayable()
    {
        using var recorded = new TempDirectory();
        var (exchanges, _) = await ShadowTests.Record(recorded);
        using var folder = new TempDirectory();
        var sink = new FolderSink(folder);
        var sent = sink.Send(Items(exchanges, new(long.MaxValue)), DateTime.UtcNow, 0, false);
        await Assert.That(sent.Sent).IsTrue();

        var read = Exchange.ReadFolder(folder);
        await Assert.That(read.Count).IsEqualTo(exchanges.Count);
        await Assert.That(read.Select(_ => _.Request.Body.Length)).IsEquivalentTo(exchanges.Select(_ => _.Request.Body.Length));
        await Assert.That(read[0].User).IsEqualTo(exchanges[0].User);
    }

    /// <summary>
    /// The batches of the module as the shadow reads them: a recorded session in the module's batches, gzip, with a
    /// tombstone, posted to the ingest of a shadow, and replayed the same.
    /// </summary>
    [Test]
    public async Task Batches_ReadByTheShadow()
    {
        using var recorded = new TempDirectory();
        var (exchanges, file) = await ShadowTests.Record(recorded);
        await using var shadow = await ShadowTests.StartShadow(file, _ => _.Levels[ShadowOutcome.TapDropped] = LogLevel.Information);
        var items = Items(exchanges, new(long.MaxValue));
        var last = exchanges[^1].Time!.Value.UtcDateTime;
        items.Add(TapItem.Tombstone(100, last.AddSeconds(1), "http://other/Shared%20Documents/big.pptx", "too large"));

        var boundary = Guid.NewGuid().ToString("N");
        using var request = new HttpRequestMessage(HttpMethod.Post, "/officehost-shadow/exchanges");
        request.Content = new ByteArrayContent(TapBatch.Body(items, boundary));
        request.Content.Headers.TryAddWithoutValidation("Content-Type", $"multipart/mixed; boundary=\"{boundary}\"");
        request.Content.Headers.ContentEncoding.Add("gzip");
        request.Headers.Add("X-OfficeHost-Shadow-Key", ShadowTests.Key);
        request.Headers.Add("X-Shadow-Instance", Guid.NewGuid().ToString());
        request.Headers.Add("X-Shadow-Server", "sharepoint");
        request.Headers.Add("X-Shadow-Watermark", last.AddSeconds(2).ToString("o"));
        using var response = await shadow.Client.SendAsync(request);
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.Accepted);

        var findings = await ShadowTests.Findings(shadow, 11);
        await Assert.That(findings.Count(_ => _.Contains("ShadowProcessor: Same: "))).IsEqualTo(10);
        await Assert.That(findings).Contains(_ => _.Contains("TapDropped: tombstone of other/shared documents/big.pptx") && _.Contains("too large"));
    }
}
