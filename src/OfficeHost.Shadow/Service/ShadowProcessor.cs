/// <summary>
/// What the shadow knows of a document besides the engine's state (<see cref="ShadowDocumentState"/>): when it last saw
/// it, the instances that sent it, and whether its answers can be compared. A document is untrusted for its content
/// (its Cells are not compared) when OfficeHost does not have the file SharePoint has, as after a save of a client that
/// knew it from Office's cache; for all its state when exchanges of it are missing or came out of order.
/// </summary>
sealed class Tracked(ShadowDocument document)
{
    public ShadowDocument Document { get; } = document;
    public DateTimeOffset LastSeen { get; set; }
    public DateTimeOffset LastStarted { get; set; }
    public HashSet<Guid> Instances { get; } = [];
    public string? Untrusted { get; private set; }
    public bool All { get; private set; }

    public void Untrust(string reason, bool all)
    {
        if (Untrusted == null ||
            (all && !All))
        {
            Untrusted = reason;
        }

        All |= all;
    }
}

/// <summary>
/// Replays what the <see cref="ShadowInbox"/> releases, one exchange at a time (docs/shadow.md): starts each document
/// from its content when first seen, replays with <see cref="ShadowReplay"/> to the application's OfficeHost in process,
/// and logs each finding at the level of its outcome. Documents unseen for <see cref="OfficeHostShadowOptions.IdleTimeout"/>
/// are forgotten, and start afresh at their next session.
/// </summary>
sealed class ShadowProcessor
{
    IShadowDocuments documents;
    IOfficeStorage storage;
    IFssHttpStore store;
    IOfficeLockStore locks;
    IServiceProvider services;
    ShadowInbox inbox;
    ServiceFiles files;
    OfficeHostShadowOptions options;
    ILogger<ShadowProcessor> log;
    ShadowReplay replay;
    Dictionary<string, Tracked> tracked = [];
    Dictionary<(ShadowOutcome Outcome, string Label), int> counts = [];
    DateTimeOffset summarized = DateTimeOffset.UtcNow;
    string? recordingFolder;
    int recorded;

    public ShadowProcessor(
        ShadowInbox inbox,
        ServiceFiles files,
        InProcessTarget target,
        IShadowDocuments documents,
        IOfficeStorage storage,
        IFssHttpStore store,
        IOfficeLockStore locks,
        IServiceProvider services,
        IOptions<OfficeHostShadowOptions> options,
        ILogger<ShadowProcessor> log)
    {
        this.inbox = inbox;
        this.files = files;
        this.documents = documents;
        this.storage = storage;
        this.store = store;
        this.locks = locks;
        this.services = services;
        this.options = options.Value;
        this.log = log;
        replay = new(target, files, this.options.KnownDifferences);
    }

    /// <summary>
    /// The documents the shadow knows of now.
    /// </summary>
    public int Documents => tracked.Count;

    /// <summary>
    /// Replays what the inbox releases, then reports losses, forgets idle documents, and logs the summary when due.
    /// </summary>
    public async Task Run(Cancel cancel)
    {
        foreach (var item in inbox.Release())
        {
            try
            {
                await Process(item, cancel);
            }
            catch (Exception exception) when (!cancel.IsCancellationRequested)
            {
                Log(ShadowOutcome.ShadowFailed, "exchange", null, item, [], $"{exception.GetType().Name}: {exception.Message}");
            }
        }

        foreach (var loss in inbox.Losses())
        {
            foreach (var document in tracked.Values.Where(_ => _.Instances.Contains(loss.Instance)))
            {
                document.Untrust($"{loss.Server} {loss.Reason}", true);
            }

            Count(ShadowOutcome.TapDropped, "server");
            Write(ShadowOutcome.TapDropped, "{Outcome}: {Server} {Reason}; the documents it sent are not compared until their next session", ShadowOutcome.TapDropped, loss.Server, loss.Reason);
        }

        Forget(DateTimeOffset.UtcNow);
        Summarize(DateTimeOffset.UtcNow);
    }

    async Task Process(ShadowItem item, Cancel cancel)
    {
        if (item.Exchange is not { } exchange)
        {
            if (item.Tombstone is { } missing)
            {
                Track(missing).Untrust($"an exchange missing: {item.Reason}", true);
            }

            Log(ShadowOutcome.TapDropped, "tombstone", item.Tombstone?.Key, item, [], item.Reason);
            return;
        }

        Record(exchange);
        if (ShadowComparer.Kind(exchange.Request) == "cellstorage" &&
            SoapBody.TryRead(exchange.Request) is { } body)
        {
            foreach (var (document, makes) in ShadowComparer.Documents(body))
            {
                if (tracked.TryGetValue(document.Key, out var known) &&
                    item.Started < known.LastStarted)
                {
                    // later than an exchange of it already replayed: its state is no longer what SharePoint's was
                    known.Untrust("an exchange came after a later one", true);
                    Log(ShadowOutcome.NotCompared, "cell storage", document.Key, item, [], "came after a later exchange of the document: not replayed");
                    return;
                }

                var state = known ?? await Seed(document, makes, exchange, ShadowComparer.Opens(body), item, cancel);
                state.LastSeen = DateTimeOffset.UtcNow;
                state.LastStarted = item.Started;
                state.Instances.Add(item.Instance);
            }
        }

        var replayed = await replay.Replay(exchange, cancel);
        if (replayed.NotReplayed is { } reason)
        {
            Log(ShadowOutcome.NotCompared, ShadowComparer.Kind(exchange.Request) ?? "request", null, item, [], $"not replayed: {reason}");
            return;
        }

        foreach (var finding in replayed.Findings)
        {
            var state = finding.Document == null ? null : tracked.GetValueOrDefault(finding.Document);
            var outcome = finding.Outcome;
            var why = finding.Reason;
            // a failure of OfficeHost itself (a server error, an exception) is one however the document stands
            var server = finding.Differences.Any(_ => _.Facet == "HTTP") || finding.Label == "request";
            if (state?.Untrusted is { } untrusted &&
                (state.All || finding.Label.StartsWith("Cell", StringComparison.Ordinal)) &&
                outcome is ShadowOutcome.Different or ShadowOutcome.OfficeHostFailed &&
                !server)
            {
                outcome = ShadowOutcome.NotCompared;
                why = untrusted;
            }

            Log(outcome, finding.Label, finding.Document, item, finding.Differences, why);
            if (state != null &&
                finding.Label.StartsWith("Cell PutChanges", StringComparison.Ordinal) &&
                finding.Label.EndsWith($"({ReplayCells.FromCache})", StringComparison.Ordinal))
            {
                state.Untrust("a save of a client that knew the document from Office's cache: OfficeHost has not the file SharePoint has", false);
            }
        }
    }

    Tracked Track(ShadowDocument document)
    {
        if (!tracked.TryGetValue(document.Key, out var state))
        {
            state = new(document)
            {
                LastSeen = DateTimeOffset.UtcNow
            };
            tracked[document.Key] = state;
        }

        return state;
    }

    /// <summary>
    /// Starts a document the shadow sees for the first time: writes its content to the storage, from a content query of
    /// the exchange that SharePoint answered in full (the document as it was, exactly), else from
    /// <see cref="IShadowContent"/>, and clears OfficeHost's state of it (sessions, cells, lock). A document the
    /// exchange makes (a Save As) is left for OfficeHost to make.
    /// </summary>
    async Task<Tracked> Seed(ShadowDocument document, bool makes, Exchange exchange, bool opens, ShadowItem item, Cancel cancel)
    {
        replay.Forget(document.Key);
        files.Forget(document.Key);
        var state = Track(document);
        var fileId = await documents.Find(document, cancel);
        var captured = ShadowComparer.Contents(exchange, document.Key).FirstOrDefault();
        if (captured == null &&
            makes &&
            fileId == null)
        {
            return state;
        }

        ShadowContent? content = null;
        if (captured != null)
        {
            content = new(captured, ShadowComparer.EtagVersion(exchange));
        }
        else if (services.GetService<IShadowContent>() is { } source)
        {
            try
            {
                content = await source.Get(document, cancel);
            }
            catch (Exception exception) when (!cancel.IsCancellationRequested)
            {
                Log(ShadowOutcome.ShadowFailed, "content", document.Key, item, [], $"the content from SharePoint: {exception.GetType().Name}: {exception.Message}");
            }
        }

        if (content == null)
        {
            if (fileId == null)
            {
                state.Untrust("no content of the document: none in the traffic, nor from SharePoint", true);
                return state;
            }

            state.Untrust("the content of the document as SharePoint has it is not known: the shadow's file may be older", false);
        }
        else
        {
            using var stream = new MemoryStream(content.Content, false);
            if (fileId == null)
            {
                fileId = await documents.Add(document, stream, cancel);
            }
            else
            {
                await storage.Write(fileId, stream, cancel);
            }

            if (captured == null &&
                content.Version is { } version &&
                ShadowComparer.EtagVersion(exchange) is { } traffic &&
                version != traffic)
            {
                state.Untrust($"SharePoint has version {version} of the document, the traffic version {traffic}", false);
            }
        }

        // what OfficeHost knew of the document before: sessions, cells, lock
        await store.Delete(fileId, cancel);
        if (await locks.Get(fileId, cancel) is { } existing)
        {
            await locks.CompareExchange(fileId, existing.Id, null, cancel);
        }

        if (!opens)
        {
            state.Untrust("first seen within a session: its locks and coauthoring are not known", true);
        }

        return state;
    }

    /// <summary>
    /// Forgets the documents unseen for the idle timeout, and the least recently seen beyond the most kept.
    /// </summary>
    void Forget(DateTimeOffset now)
    {
        var idle = tracked.Values.Where(_ => now - _.LastSeen > options.IdleTimeout);
        var excess = tracked.Values
            .OrderBy(_ => _.LastSeen)
            .Take(Math.Max(0, tracked.Count - options.MaximumDocuments));
        foreach (var document in idle.Concat(excess).Distinct().ToList())
        {
            tracked.Remove(document.Document.Key);
            replay.Forget(document.Document.Key);
            files.Forget(document.Document.Key);
        }
    }

    void Record(Exchange exchange)
    {
        if (options.RecordingFolder is not { } root)
        {
            return;
        }

        var day = DateTimeOffset.UtcNow.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
        var folder = Path.Combine(root, day);
        if (recordingFolder != folder)
        {
            Directory.CreateDirectory(folder);
            recordingFolder = folder;
            recorded = Directory.EnumerateFiles(folder, "*.request.http").Count();
        }

        recorded++;
        (exchange with { Number = recorded }).Write(folder);
    }

    void Count(ShadowOutcome outcome, string label) =>
        counts[(outcome, label)] = counts.GetValueOrDefault((outcome, label)) + 1;

    void Log(ShadowOutcome outcome, string label, string? document, ShadowItem item, IReadOnlyList<ShadowDifference> differences, string? reason)
    {
        Count(outcome, label);
        var detail = "";
        if (differences.Count > 0)
        {
            detail = $": {string.Join("; ", differences)}";
        }

        if (reason != null)
        {
            detail += $" ({reason})";
        }

        Write(
            outcome,
            "{Outcome}: {Label} of {Document} by {User}, exchange {Sequence} of {Server} at {Started:O}{Detail}",
            outcome,
            label,
            document ?? "(none)",
            item.Exchange?.User ?? "",
            item.Sequence,
            item.Server,
            item.Started,
            detail);
    }

    void Write(ShadowOutcome outcome, string message, params object?[] arguments)
    {
        var level = options.Levels.GetValueOrDefault(outcome, LogLevel.Information);
        if (log.IsEnabled(level))
        {
#pragma warning disable CA2254 // the template is one of this class's
            log.Log(level, new EventId((int) outcome, outcome.ToString()), message, arguments);
#pragma warning restore CA2254
        }
    }

    /// <summary>
    /// The counts of the outcomes since the last summary, by label, at Information.
    /// </summary>
    void Summarize(DateTimeOffset now)
    {
        if (now - summarized < options.SummaryInterval)
        {
            return;
        }

        var (received, refused) = inbox.Totals;
        var outcomes = string.Join(
            "; ",
            counts.OrderBy(_ => _.Key.Outcome).ThenBy(_ => _.Key.Label, StringComparer.Ordinal).Select(_ => $"{_.Key.Outcome} {_.Key.Label} x{_.Value}"));
        log.LogInformation(
            new EventId(100, "Summary"),
            "Shadow since {Since:O}: {Outcomes}. {Documents} documents known, {Waiting} exchanges waiting, {Received} received and {Refused} refused since the start",
            summarized,
            outcomes,
            tracked.Count,
            inbox.Count,
            received,
            refused);
        counts.Clear();
        summarized = now;
    }
}
