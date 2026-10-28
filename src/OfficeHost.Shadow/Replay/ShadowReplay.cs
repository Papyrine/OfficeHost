/// <summary>
/// The OfficeHost the shadow replays to: in process (<see cref="InProcessTarget"/>), or a test server.
/// </summary>
interface IShadowTarget
{
    /// <summary>
    /// Sends a request to a path of OfficeHost's desktop site, as at the time it was recorded.
    /// </summary>
    Task<HttpMessage> Send(string path, HttpMessage recorded, byte[] body, DateTimeOffset? time, Cancel cancel);
}

/// <summary>
/// A file of the shadow's storage, and the path of its desktop site for a user (<c>/fsshttp/{token}</c>).
/// </summary>
record ShadowFile(string FileId, string Site);

/// <summary>
/// The files of the documents the shadow replays, for a user of SharePoint (as it named them).
/// </summary>
interface IShadowFiles
{
    /// <summary>
    /// The file of a document for a user, or null when the storage has none.
    /// </summary>
    Task<ShadowFile?> Open(ShadowDocument document, string user, Cancel cancel);
}

/// <summary>
/// What the shadow knows of a document of SharePoint: the ids SharePoint made for it, translated to OfficeHost's, and
/// the copies of its cells (<see cref="ReplayCells"/>).
/// </summary>
class ShadowDocumentState(ShadowDocument document)
{
    public ShadowDocument Document { get; } = document;

    // the GUIDs of SharePoint to those of OfficeHost: its ResourceID and TransitionID
    public Dictionary<Guid, Guid> Ids { get; } = [];

    public Guid? ResourceId { get; set; }

    public ReplayCells Cells { get; set; } = new();
}

/// <summary>
/// The result of the replay of an exchange: its findings, or why it was not replayed.
/// </summary>
record ShadowReplayed(List<ShadowFinding> Findings, string? NotReplayed = null);

/// <summary>
/// Replays exchanges recorded with SharePoint to OfficeHost and compares the answers (docs/replay.md, docs/shadow.md).
/// The cell storage and shared access requests are replayed, each as its user, at its time. The ids SharePoint made
/// are translated to OfficeHost's: the ResourceID and the TransitionID, learned from the pairs of answers, and those of
/// the cells by <see cref="ReplayCells"/>. A Save As, to a name the storage has no file of, goes through the desktop
/// site of the last document of the same user in the same folder, as Office does.
/// </summary>
class ShadowReplay(IShadowTarget target, IShadowFiles files, IReadOnlyList<ShadowRule> rules)
{
    static XNamespace soap = SoapBody.Soap;

    static Regex guids = new(@"\{?[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}\}?");

    Dictionary<string, ShadowDocumentState> documents = [];

    // the ResourceID of SharePoint to the key of the document, as IsOnlyClient names it
    Dictionary<Guid, string> resources = [];

    // the last file of each user in each folder, whose site a Save As goes through
    Dictionary<(string User, string Folder), ShadowFile> sites = [];

    public ShadowDocumentState State(ShadowDocument document)
    {
        if (!documents.TryGetValue(document.Key, out var state))
        {
            state = new(document);
            documents[document.Key] = state;
        }

        return state;
    }

    /// <summary>
    /// Forgets a document: its ids and the copies of its cells, as when it is evicted or written again.
    /// </summary>
    public void Forget(string key)
    {
        if (!documents.Remove(key, out var state))
        {
            return;
        }

        if (state.ResourceId is { } resourceId)
        {
            resources.Remove(resourceId);
        }

        foreach (var site in sites.Where(_ => _.Key.Folder == state.Document.Folder).Select(_ => _.Key).ToList())
        {
            sites.Remove(site);
        }
    }

    public async Task<ShadowReplayed> Replay(Exchange exchange, Cancel cancel = default)
    {
        switch (ShadowComparer.Kind(exchange.Request))
        {
            case "cellstorage":
                return await ReplayCellStorage(exchange, cancel);
            case "sharedaccess":
                return await ReplaySharedAccess(exchange, cancel);
            default:
                return new([], "not a request of the cell storage or the shared access service");
        }
    }

    async Task<ShadowReplayed> ReplayCellStorage(Exchange exchange, Cancel cancel)
    {
        if (SoapBody.TryRead(exchange.Request) is not { } body)
        {
            return new([], $"no SOAP message in the recorded request ({exchange.Request.Body.Length} bytes), SharePoint answered {exchange.Response.Status}");
        }

        var named = ShadowComparer.Documents(body).Select(_ => State(_.Document)).ToList();
        var document = Translate(body.Document, named);
        var requests = document.Descendants(soap + "Request").ToList();
        var translations = new Dictionary<(int Request, string? Token), ReplayCells.Translation>();
        string? site = null;
        for (var index = 0; index < requests.Count; index++)
        {
            var request = requests[index];
            if (ShadowDocument.From((string?) request.Attribute("Url")) is not { } shadowDocument)
            {
                break;
            }

            var file = await files.Open(shadowDocument, exchange.User, cancel);
            if (file != null)
            {
                sites[(exchange.User, shadowDocument.Folder)] = file;
            }
            else if (!sites.TryGetValue((exchange.User, shadowDocument.Folder), out file))
            {
                break;
            }

            site ??= file.Site;
            request.SetAttributeValue("Url", $"http://localhost{file.Site}/{Uri.EscapeDataString(shadowDocument.Name)}");
            var state = State(shadowDocument);
            foreach (var subRequest in request.Elements(soap + "SubRequest").Where(_ => (string?) _.Attribute("Type") == "Cell"))
            {
                var data = subRequest.Element(soap + "SubRequestData");
                if (ShadowComparer.RequestOf(body.Binary(data)) is not { } cell)
                {
                    continue;
                }

                var bytes = state.Cells.Translate(shadowDocument.Key, ShadowComparer.PartitionOf(data), cell, out var translation);
                translations[(index, (string?) subRequest.Attribute("SubRequestToken"))] = translation;
                if (bytes != null)
                {
                    // inline, as base64: OfficeHost reads either
                    data!.RemoveNodes();
                    data.Add(Convert.ToBase64String(bytes));
                    data.SetAttributeValue("BinaryDataSize", bytes.Length);
                }
            }
        }

        if (site == null)
        {
            return new([], "no file of the document in the storage, nor a document of the user in its folder to make it from");
        }

        var response = await Send($"{site}/_vti_bin/cellstorage.svc/CellStorageService", exchange, body.With(document), cancel);
        if (response.Failed is { } failed)
        {
            return new([failed]);
        }

        return new(Compare(exchange, body, response.Message!, translations));
    }

    async Task<ShadowReplayed> ReplaySharedAccess(Exchange exchange, Cancel cancel)
    {
        XDocument recorded;
        try
        {
            recorded = XDocument.Parse(Encoding.UTF8.GetString(exchange.Request.Body));
        }
        catch (XmlException)
        {
            return new([], "no SOAP message in the recorded request");
        }

        var id = recorded.Descendants().FirstOrDefault(_ => _.Name.LocalName == "id")?.Value;
        if (!Guid.TryParse(id, out var resourceId) ||
            !resources.TryGetValue(resourceId, out var key) ||
            !documents.TryGetValue(key, out var state))
        {
            return new([], "no document of its id seen in a cell storage answer");
        }

        var file = await files.Open(state.Document, exchange.User, cancel);
        if (file == null)
        {
            return new([], "no file of the document in the storage");
        }

        var response = await Send($"{file.Site}/_vti_bin/sharedaccess.asmx", exchange, Encoding.UTF8.GetBytes(Translate(recorded, [state]).ToString(SaveOptions.DisableFormatting)), cancel);
        if (response.Failed is { } failed)
        {
            return new([failed]);
        }

        var message = response.Message!;
        const string label = "IsOnlyClient";
        if (message.Status != exchange.Response.Status)
        {
            return new([Http(exchange, key, label, message)]);
        }

        var expected = ShadowComparer.IsOnlyClientResult(exchange.Response);
        var actual = ShadowComparer.IsOnlyClientResult(message);
        if (expected == actual)
        {
            return new([new(exchange, key, label, ShadowOutcome.Same, [], expected)]);
        }

        return new([new(exchange, key, label, ShadowOutcome.Different, [new("IsOnlyClientResult", expected, actual)])]);
    }

    /// <summary>
    /// Sends the request to OfficeHost: its answer, or a finding of OfficeHost failing when it threw.
    /// </summary>
    async Task<(HttpMessage? Message, ShadowFinding? Failed)> Send(string path, Exchange exchange, byte[] body, Cancel cancel)
    {
        try
        {
            return (await target.Send(path, exchange.Request, body, exchange.Time, cancel), null);
        }
        catch (Exception exception) when (!cancel.IsCancellationRequested)
        {
            return (null, new(exchange, null, "request", ShadowOutcome.OfficeHostFailed, [], Reason: $"{exception.GetType().Name}: {exception.Message}"));
        }
    }

    static ShadowFinding Http(Exchange exchange, string? document, string label, HttpMessage response)
    {
        var outcome = ShadowOutcome.Different;
        if (response.Status >= 500)
        {
            outcome = ShadowOutcome.OfficeHostFailed;
        }

        return new(exchange, document, label, outcome, [new("HTTP", $"{exchange.Response.Status}", $"{response.Status}")]);
    }

    /// <summary>
    /// The request with the GUIDs of SharePoint that OfficeHost has its own of replaced by those, written as they were
    /// (braces, case).
    /// </summary>
    static XDocument Translate(XDocument document, List<ShadowDocumentState> states)
    {
        var text = guids.Replace(
            document.ToString(SaveOptions.DisableFormatting),
            match =>
            {
                if (!Guid.TryParse(match.Value, out var recorded))
                {
                    return match.Value;
                }

                var replay = states.Select(_ => _.Ids.GetValueOrDefault(recorded)).FirstOrDefault(_ => _ != Guid.Empty);
                if (replay == Guid.Empty)
                {
                    return match.Value;
                }

                var value = replay.ToString();
                if (match.Value.Any(char.IsUpper))
                {
                    value = value.ToUpperInvariant();
                }

                if (match.Value.StartsWith('{'))
                {
                    return $"{{{value}}}";
                }

                return value;
            });
        return XDocument.Parse(text);
    }

    static void Learn(ShadowDocumentState state, string? recorded, string? replay)
    {
        if (Guid.TryParse(recorded, out var from) &&
            Guid.TryParse(replay, out var to))
        {
            state.Ids[from] = to;
        }
    }

    List<ShadowFinding> Compare(Exchange exchange, SoapBody request, HttpMessage response, Dictionary<(int Request, string? Token), ReplayCells.Translation> translations)
    {
        var requests = request.Document.Descendants(soap + "Request").ToList();
        var first = ShadowDocument.From((string?) requests.FirstOrDefault()?.Attribute("Url"))?.Key;
        if (exchange.Response.Status != 200 ||
            response.Status != 200)
        {
            if (exchange.Response.Status == response.Status)
            {
                return [new(exchange, first, "cell storage", ShadowOutcome.Same, [], $"HTTP {response.Status}")];
            }

            return [Http(exchange, first, "cell storage", response)];
        }

        var findings = new List<ShadowFinding>();
        var expected = SoapBody.Read(exchange.Response);
        var actual = SoapBody.Read(response);
        var expectedResponses = expected.Document.Descendants(soap + "Response").ToList();
        var actualResponses = actual.Document.Descendants(soap + "Response").ToList();
        for (var index = 0; index < expectedResponses.Count; index++)
        {
            var shadowDocument = ShadowDocument.From((string?) requests.ElementAtOrDefault(index)?.Attribute("Url"));
            var key = shadowDocument?.Key;
            var expectedResponse = expectedResponses[index];
            var actualResponse = actualResponses.ElementAtOrDefault(index);
            if (actualResponse == null)
            {
                findings.Add(new(exchange, key, "cell storage", ShadowOutcome.Different, [new("Response", "(present)", null)]));
                continue;
            }

            var state = shadowDocument == null ? null : State(shadowDocument);
            if (state != null)
            {
                Learn(state, (string?) expectedResponse.Attribute("ResourceID"), (string?) actualResponse.Attribute("ResourceID"));
                if (Guid.TryParse((string?) expectedResponse.Attribute("ResourceID"), out var resourceId))
                {
                    state.ResourceId = resourceId;
                    resources[resourceId] = state.Document.Key;
                }
            }

            var subRequests = requests.ElementAtOrDefault(index)?.Elements(soap + "SubRequest").ToList() ?? [];
            var actualSubResponses = actualResponse.Elements(soap + "SubResponse").ToList();
            foreach (var expectedSubResponse in expectedResponse.Elements(soap + "SubResponse"))
            {
                var token = (string?) expectedSubResponse.Attribute("SubRequestToken");
                var subRequest = subRequests.FirstOrDefault(_ => (string?) _.Attribute("SubRequestToken") == token);
                var translation = translations.GetValueOrDefault((index, token));
                var label = ShadowComparer.Label(subRequest, request, translation?.Status);
                var actualSubResponse = actualSubResponses.FirstOrDefault(_ => (string?) _.Attribute("SubRequestToken") == token);
                if (actualSubResponse == null)
                {
                    findings.Add(new(exchange, key, label, ShadowOutcome.Different, [new("SubResponse", "(present)", null)]));
                    continue;
                }

                var expectedData = expectedSubResponse.Element(soap + "SubResponseData");
                var actualData = actualSubResponse.Element(soap + "SubResponseData");
                if (state != null)
                {
                    Learn(state, (string?) expectedData?.Attribute("TransitionID"), (string?) actualData?.Attribute("TransitionID"));
                }

                var expectedFacets = ShadowComparer.Facets(expectedSubResponse, expected);
                var actualFacets = ShadowComparer.Facets(actualSubResponse, actual);
                var contentCompared = false;
                if (translation != null &&
                    state != null)
                {
                    // the copies of the partition after the answers, when both copies have it all: what each server
                    // has sent its client so far, rather than this answer alone, which holds only what was new to it
                    state.Cells.Answered(translation, ShadowComparer.ResponseOf(expected.Binary(expectedData)), ShadowComparer.ResponseOf(actual.Binary(actualData)));
                    if (state.Cells.File(translation.Document, translation.Partition, true) is { } recordedFile &&
                        state.Cells.File(translation.Document, translation.Partition, false) is { } replayFile)
                    {
                        foreach (var facets in new[] { expectedFacets, actualFacets })
                        {
                            facets.Remove("file");
                            facets.Remove("editors");
                        }

                        if (ShadowComparer.EditorsOf(recordedFile) is { } recordedEditors &&
                            ShadowComparer.EditorsOf(replayFile) is { } replayEditors)
                        {
                            expectedFacets["editors"] = recordedEditors;
                            actualFacets["editors"] = replayEditors;
                        }
                        else
                        {
                            expectedFacets["content"] = ShadowComparer.Hash(recordedFile);
                            actualFacets["content"] = ShadowComparer.Hash(replayFile);
                        }

                        contentCompared = true;
                    }
                }

                var differences = ShadowComparer.Differences(expectedFacets, actualFacets);
                var outcome = ShadowComparer.Outcome(label, differences, rules, out var reason);
                (byte[], byte[])? differentFiles = null;
                if (outcome != ShadowOutcome.Same &&
                    ShadowComparer.FileOf(expected.Binary(expectedData)) is { } expectedFile &&
                    ShadowComparer.FileOf(actual.Binary(actualData)) is { } actualFile &&
                    !expectedFile.AsSpan().SequenceEqual(actualFile))
                {
                    differentFiles = (expectedFile, actualFile);
                }

                findings.Add(new(exchange, key, label, outcome, differences, expectedFacets.GetValueOrDefault("ErrorCode"), reason, contentCompared, differentFiles));
            }
        }

        return findings;
    }
}
