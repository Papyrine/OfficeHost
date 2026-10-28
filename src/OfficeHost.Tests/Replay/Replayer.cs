using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;

/// <summary>
/// What a replay found: a line per exchange or sub-response, and the sub-responses that differ by their label.
/// </summary>
record ReplayResult(List<string> Lines, int Same, Dictionary<string, int> Different, int Skipped)
{
    public override string ToString() => string.Join(Environment.NewLine, Lines);
}

/// <summary>
/// Replays exchanges recorded with a server of Office documents, SharePoint, to an OfficeHost of its own, and compares
/// the answers (docs/replay.md). The cell storage and shared access requests are replayed in order, each as its
/// recorded user, at its recorded time; the other requests are not. The documents start as the first full content
/// query of the recording answered them, or as a file of the same name in the files folder.
///
/// The ids the recorded server made are translated in the requests to those OfficeHost made: the ResourceID and the
/// TransitionID, learned from the pairs of answers. Those of the cells are not: a Put Changes, or a query with the
/// knowledge of the recorded server, carries ids OfficeHost does not have, and its Cell is labelled "untranslated".
/// </summary>
class Replayer : IAsyncDisposable
{
    static XNamespace soap = SoapBody.Soap;

    // the attributes whose values each server makes its own: compared by whether they are there
    static HashSet<string> volatileAttributes =
    [
        "TransitionID", "ResourceID", "Etag", "CreateTime", "LastModifiedTime", "ModifiedBy", "ServerTime", "UserName",
        "UserLogin", "UserEmailAddress", "UserSIPAddress", "HealthScore", "Url", "WebUrl"
    ];

    static Regex guids = new(@"\{?[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}\}?");

    TestServerHost host = null!;
    ManualTime time = new();

    // the GUIDs of the recorded server to those of OfficeHost
    Dictionary<Guid, Guid> ids = [];

    // the ResourceID of the recorded server to the name of the document
    Dictionary<Guid, string> documents = [];
    Dictionary<(string Document, string User), string> tokens = [];
    string? lastDocument;

    List<string> lines = [];
    int same;
    Dictionary<string, int> different = [];
    int skipped;

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
        foreach (var exchange in exchanges)
        {
            await replayer.Replay(exchange);
        }

        var differences = replayer.different.Values.Sum();
        replayer.lines.Add($"{replayer.same} the same, {differences} different, {replayer.skipped} requests not replayed");
        foreach (var (label, count) in replayer.different.OrderBy(_ => _.Key))
        {
            replayer.lines.Add($"  {label}: {count} different");
        }

        return new(replayer.lines, replayer.same, replayer.different, replayer.skipped);
    }

    /// <summary>
    /// Writes each document of the cell storage requests to the storage: the file of its name in the files folder, or
    /// the file of the first content query the recorded server answered in full.
    /// </summary>
    async Task Seed(List<Exchange> exchanges, string? files)
    {
        var names = new List<string>();
        foreach (var exchange in exchanges.Where(_ => Kind(_.Request) == "cellstorage"))
        {
            var request = SoapBody.Read(exchange.Request);
            foreach (var name in request.Document.Descendants(soap + "Request").Select(_ => DocumentName((string?) _.Attribute("Url"))))
            {
                if (name != null &&
                    !names.Contains(name))
                {
                    names.Add(name);
                }
            }
        }

        foreach (var name in names)
        {
            byte[]? content = null;
            if (files != null &&
                File.Exists(Path.Combine(files, name)))
            {
                content = await File.ReadAllBytesAsync(Path.Combine(files, name));
            }

            content ??= exchanges
                .Where(_ => Kind(_.Request) == "cellstorage" && _.Response.Status == 200)
                .SelectMany(_ => Contents(_, name))
                .FirstOrDefault();
            if (content == null)
            {
                lines.Add($"No content for {name}: not seeded (put the file in the files folder to seed it)");
                continue;
            }

            await File.WriteAllBytesAsync(host.PathOf(name), content);
            lines.Add($"Seeded {name}, {content.Length} bytes");
        }
    }

    /// <summary>
    /// The files that the content queries of a recorded exchange for a document answer in full.
    /// </summary>
    static IEnumerable<byte[]> Contents(Exchange exchange, string name)
    {
        var request = SoapBody.Read(exchange.Request);
        var response = SoapBody.Read(exchange.Response);
        var urls = request.Document.Descendants(soap + "Request")
            .Select(_ => ((string?) _.Attribute("RequestToken"), DocumentName((string?) _.Attribute("Url"))))
            .ToList();
        foreach (var item in response.Document.Descendants(soap + "Response"))
        {
            if (!urls.Contains(((string?) item.Attribute("RequestToken"), name)))
            {
                continue;
            }

            foreach (var data in item.Elements(soap + "SubResponse").Select(_ => _.Element(soap + "SubResponseData")))
            {
                if (FileOf(response.Binary(data)) is { } file)
                {
                    yield return file;
                }
            }
        }
    }

    /// <summary>
    /// The file a Cell answer holds in full, or null.
    /// </summary>
    static byte[]? FileOf(byte[]? binary)
    {
        if (binary == null)
        {
            return null;
        }

        try
        {
            var response = FssHttpBResponse.Read(binary);
            if (response.SubResponses.OfType<QueryChangesResponse>().FirstOrDefault() is not { } query)
            {
                return null;
            }

            var storage = CellStorage.Create(response.DataElements?.Elements ?? [], query.StorageIndex);
            return FileCell.Read(storage).Root.ToArray();
        }
        catch (Exception exception) when (exception is FssHttpBException or FssHttpDException or InvalidDataException or KeyNotFoundException)
        {
            return null;
        }
    }

    static string? Kind(HttpMessage request)
    {
        var path = request.Target.Split('?')[0];
        if (path.EndsWith("/_vti_bin/cellstorage.svc/CellStorageService", StringComparison.OrdinalIgnoreCase))
        {
            return "cellstorage";
        }

        if (path.EndsWith("/_vti_bin/sharedaccess.asmx", StringComparison.OrdinalIgnoreCase))
        {
            return "sharedaccess";
        }

        return null;
    }

    /// <summary>
    /// The name of the document a Request Url names: its last segment.
    /// </summary>
    static string? DocumentName(string? url)
    {
        if (url == null ||
            !Uri.TryCreate(url, UriKind.Absolute, out var uri))
        {
            return null;
        }

        var name = Uri.UnescapeDataString(uri.AbsolutePath.TrimEnd('/').Split('/')[^1]);
        if (name.Length == 0)
        {
            return null;
        }

        return name;
    }

    /// <summary>
    /// The user of OfficeHost of a user of the recorded server: the name without its domain or claim prefix
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

    /// <summary>
    /// The token of a document for a user: of the document when the storage has it, else of the last document, whose
    /// site names the new one as a Save As does.
    /// </summary>
    async Task<string?> Token(string? document, string user)
    {
        if (document != null &&
            File.Exists(host.PathOf(document)))
        {
            lastDocument = document;
        }
        else
        {
            document = lastDocument;
        }

        if (document == null)
        {
            return null;
        }

        if (!tokens.TryGetValue((document, user), out var token))
        {
            (_, token) = await host.Open(document, userId: user);
            tokens[(document, user)] = token;
        }

        return token;
    }

    async Task Replay(Exchange exchange)
    {
        if (exchange.Time is { } at)
        {
            time.Now = at;
        }

        switch (Kind(exchange.Request))
        {
            case "cellstorage":
                await ReplayCellStorage(exchange);
                break;
            case "sharedaccess":
                await ReplaySharedAccess(exchange);
                break;
            default:
                skipped++;
                break;
        }
    }

    async Task ReplayCellStorage(Exchange exchange)
    {
        var user = UserOf(exchange.User);
        var body = SoapBody.Read(exchange.Request);
        var document = Translate(body.Document);
        var requests = document.Descendants(soap + "Request").ToList();
        string? site = null;
        foreach (var request in requests)
        {
            var name = DocumentName((string?) request.Attribute("Url"));
            var token = await Token(name, user);
            if (token == null)
            {
                break;
            }

            site ??= token;
            request.SetAttributeValue("Url", $"http://localhost/fsshttp/{token}/{Uri.EscapeDataString(name ?? "")}");
        }

        if (site == null)
        {
            lines.Add($"#{exchange.Number} cell storage: no document to replay it to");
            skipped++;
            return;
        }

        var response = await Send($"/fsshttp/{site}/_vti_bin/cellstorage.svc/CellStorageService", exchange.Request, body.With(document));
        await Compare(exchange, body, response);
    }

    async Task ReplaySharedAccess(Exchange exchange)
    {
        var user = UserOf(exchange.User);
        var recorded = XDocument.Parse(Encoding.UTF8.GetString(exchange.Request.Body));
        var id = recorded.Descendants().FirstOrDefault(_ => _.Name.LocalName == "id")?.Value;
        string? name = null;
        if (Guid.TryParse(id, out var resourceId))
        {
            name = documents.GetValueOrDefault(resourceId);
        }

        var token = await Token(name, user);
        if (token == null)
        {
            lines.Add($"#{exchange.Number} IsOnlyClient: no document to replay it to");
            skipped++;
            return;
        }

        var response = await Send($"/fsshttp/{token}/_vti_bin/sharedaccess.asmx", exchange.Request, Encoding.UTF8.GetBytes(Translate(recorded).ToString(SaveOptions.DisableFormatting)));
        var label = "IsOnlyClient";
        if (response.Status != exchange.Response.Status)
        {
            Differ(exchange, label, [$"HTTP {exchange.Response.Status} / {response.Status}"]);
            return;
        }

        var expected = Result(exchange.Response);
        var actual = Result(response);
        if (expected == actual)
        {
            Same(exchange, label, $"{expected}");
            return;
        }

        Differ(exchange, label, [$"IsOnlyClientResult {expected} / {actual}"]);
    }

    static string? Result(HttpMessage response)
    {
        try
        {
            return XDocument.Parse(Encoding.UTF8.GetString(response.Body)).Descendants().FirstOrDefault(_ => _.Name.LocalName == "IsOnlyClientResult")?.Value;
        }
        catch (XmlException)
        {
            return null;
        }
    }

    async Task<HttpMessage> Send(string path, HttpMessage recorded, byte[] body)
    {
        using var request = new HttpRequestMessage(HttpMethod.Post, path);
        request.Content = new ByteArrayContent(body);
        request.Content.Headers.TryAddWithoutValidation("Content-Type", recorded.ContentType);
        if (recorded.Header("SOAPAction") is { } action)
        {
            request.Headers.TryAddWithoutValidation("SOAPAction", action);
        }

        using var response = await host.Client.SendAsync(request);
        var headers = response.Content.Headers.Select(_ => (_.Key, string.Join(", ", _.Value))).ToList();
        return new($"HTTP/1.1 {(int) response.StatusCode} {response.ReasonPhrase}", headers, await response.Content.ReadAsByteArrayAsync());
    }

    /// <summary>
    /// The request with the GUIDs of the recorded server that OfficeHost has its own of replaced by those, written as
    /// they were (braces, case).
    /// </summary>
    XDocument Translate(XDocument document)
    {
        var text = guids.Replace(
            document.ToString(SaveOptions.DisableFormatting),
            match =>
            {
                if (!Guid.TryParse(match.Value, out var recorded) ||
                    !ids.TryGetValue(recorded, out var replay))
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

    void Learn(string? recorded, string? replay)
    {
        if (Guid.TryParse(recorded, out var from) &&
            Guid.TryParse(replay, out var to))
        {
            ids[from] = to;
        }
    }

    async Task Compare(Exchange exchange, SoapBody request, HttpMessage response)
    {
        if (exchange.Response.Status != 200 ||
            response.Status != 200)
        {
            if (exchange.Response.Status == response.Status)
            {
                Same(exchange, "cell storage", $"HTTP {response.Status}");
                return;
            }

            Differ(exchange, "cell storage", [$"HTTP {exchange.Response.Status} / {response.Status}"]);
            return;
        }

        var expected = SoapBody.Read(exchange.Response);
        var actual = SoapBody.Read(response);
        var expectedResponses = expected.Document.Descendants(soap + "Response").ToList();
        var actualResponses = actual.Document.Descendants(soap + "Response").ToList();
        var requests = request.Document.Descendants(soap + "Request").ToList();
        for (var index = 0; index < expectedResponses.Count; index++)
        {
            var expectedResponse = expectedResponses[index];
            var actualResponse = actualResponses.ElementAtOrDefault(index);
            if (actualResponse == null)
            {
                Differ(exchange, "cell storage", ["a Response missing"]);
                continue;
            }

            Learn((string?) expectedResponse.Attribute("ResourceID"), (string?) actualResponse.Attribute("ResourceID"));
            if (Guid.TryParse((string?) expectedResponse.Attribute("ResourceID"), out var resourceId) &&
                DocumentName((string?) requests.ElementAtOrDefault(index)?.Attribute("Url")) is { } name)
            {
                documents[resourceId] = name;
            }

            var subRequests = requests.ElementAtOrDefault(index)?.Elements(soap + "SubRequest").ToList() ?? [];
            var actualSubResponses = actualResponse.Elements(soap + "SubResponse").ToList();
            foreach (var expectedSubResponse in expectedResponse.Elements(soap + "SubResponse"))
            {
                var token = (string?) expectedSubResponse.Attribute("SubRequestToken");
                var subRequest = subRequests.FirstOrDefault(_ => (string?) _.Attribute("SubRequestToken") == token);
                var label = Label(subRequest, request);
                var actualSubResponse = actualSubResponses.FirstOrDefault(_ => (string?) _.Attribute("SubRequestToken") == token);
                if (actualSubResponse == null)
                {
                    Differ(exchange, label, ["no SubResponse"]);
                    continue;
                }

                var expectedData = expectedSubResponse.Element(soap + "SubResponseData");
                var actualData = actualSubResponse.Element(soap + "SubResponseData");
                Learn((string?) expectedData?.Attribute("TransitionID"), (string?) actualData?.Attribute("TransitionID"));
                var expectedFacets = Facets(expectedSubResponse, expected);
                var actualFacets = Facets(actualSubResponse, actual);
                var differences = Differences(expectedFacets, actualFacets);
                if (differences.Count == 0)
                {
                    Same(exchange, label, expectedFacets.GetValueOrDefault("ErrorCode") ?? "");
                    continue;
                }

                Differ(exchange, label, differences);
                if (output != null &&
                    FileOf(expected.Binary(expectedData)) is { } expectedFile &&
                    FileOf(actual.Binary(actualData)) is { } actualFile &&
                    !expectedFile.AsSpan().SequenceEqual(actualFile))
                {
                    var folder = Directory.CreateDirectory(Path.Combine(output, "differences")).FullName;
                    var extension = Path.GetExtension(DocumentName((string?) requests.ElementAtOrDefault(index)?.Attribute("Url")) ?? "");
                    var prefix = Path.Combine(folder, $"{exchange.Number:0000}-{token}");
                    await File.WriteAllBytesAsync($"{prefix}-sharepoint{extension}", expectedFile);
                    await File.WriteAllBytesAsync($"{prefix}-officehost{extension}", actualFile);
                    lines.Add($"    files written to {prefix}-*{extension}");
                }
            }
        }
    }

    /// <summary>
    /// The sub-request type, with its request type (<c>Coauth JoinCoauthoring</c>), or the FSSHTTPB sub-requests of a
    /// Cell (<c>Cell QueryChanges</c>), "untranslated" when the ids of the recorded server are in it.
    /// </summary>
    static string Label(XElement? subRequest, SoapBody request)
    {
        if (subRequest == null)
        {
            return "?";
        }

        var type = (string?) subRequest.Attribute("Type") ?? "?";
        var data = subRequest.Element(soap + "SubRequestData");
        if (type == "Cell")
        {
            try
            {
                if (request.Binary(data) is { } binary)
                {
                    var fssHttpB = FssHttpBRequest.Read(binary);
                    var types = string.Join(" ", fssHttpB.SubRequests.Select(_ => _.Type));
                    var untranslated = fssHttpB.SubRequests.Any(_ =>
                        _ is PutChangesRequest ||
                        _ is QueryChangesRequest { Knowledge.Items.Count: > 0 });
                    if (untranslated)
                    {
                        return $"Cell {types} (untranslated)";
                    }

                    return $"Cell {types}";
                }
            }
            catch (Exception exception) when (exception is FssHttpBException or InvalidDataException or FormatException)
            {
                return "Cell (unreadable)";
            }
        }

        var requestType = data?.Attributes()
            .FirstOrDefault(_ => _.Name.LocalName is "CoauthRequestType" or "SchemaLockRequestType" or "ExclusiveLockRequestType" or "EditorsTableRequestType" or "FileOperation" or "VersioningRequestType" or "Properties");
        if (requestType == null)
        {
            return type;
        }

        return $"{type} {requestType.Value}";
    }

    /// <summary>
    /// What is compared of a SubResponse: its error code and message, the attributes of its data (the volatile ones by
    /// whether they are there), the keys of its property lists, and the FSSHTTPB of a Cell, with the hash of the file a
    /// query answers in full.
    /// </summary>
    static Dictionary<string, string> Facets(XElement subResponse, SoapBody body)
    {
        var facets = new Dictionary<string, string>
        {
            ["ErrorCode"] = (string?) subResponse.Attribute("ErrorCode") ?? ""
        };
        if ((string?) subResponse.Attribute("ErrorMessage") is { } message)
        {
            facets["ErrorMessage"] = message;
        }

        var data = subResponse.Element(soap + "SubResponseData");
        if (data == null)
        {
            return facets;
        }

        foreach (var attribute in data.Attributes().Where(_ => !_.IsNamespaceDeclaration))
        {
            var name = attribute.Name.LocalName;
            var value = attribute.Value;
            if (volatileAttributes.Contains(name))
            {
                value = "*";
            }

            facets[$"@{name}"] = value;
        }

        foreach (var child in data.Elements().Where(_ => _.Name.LocalName != "Include"))
        {
            var keys = child.Descendants()
                .Select(_ => (string?) _.Attribute("Key") ?? _.Name.LocalName)
                .Order(StringComparer.Ordinal);
            facets[child.Name.LocalName] = string.Join(",", keys);
        }

        if (body.Binary(data) is { } binary)
        {
            facets["FSSHTTPB"] = Describe(binary);
            if (FileOf(binary) is { } file)
            {
                if (EditorsOf(file) is { } editors)
                {
                    facets["editors"] = editors;
                }
                else
                {
                    facets["file"] = Convert.ToHexString(SHA256.HashData(file))[..12];
                }
            }
        }

        return facets;
    }

    // the Editors Table Zip Stream Header ([MS-FSSHTTPB] 2.2.3.1.2.1.1)
    static byte[] editorsTableHeader = [0x1A, 0x5A, 0x3A, 0x30, 0x00, 0x00, 0x00, 0x00];

    /// <summary>
    /// An editors table by what the servers agree on: each editor's CacheID, permission and metadata keys. Not its login
    /// name (the users are mapped), its timeout (OfficeHost keeps clients an hour at least), nor the layout of its XML.
    /// </summary>
    static string? EditorsOf(byte[] file)
    {
        if (!file.AsSpan().StartsWith(editorsTableHeader))
        {
            return null;
        }

        using var deflate = new DeflateStream(new MemoryStream(file, editorsTableHeader.Length, file.Length - editorsTableHeader.Length), CompressionMode.Decompress);
        var table = XDocument.Load(deflate);
        return string.Join(
            "; ",
            table.Root!.Elements("Editor")
                .Select(_ =>
                {
                    var keys = (_.Element("Metadata")?.Elements() ?? []).Select(_ => _.Name.LocalName).Order(StringComparer.Ordinal);
                    return $"{((string?) _.Element("CacheID"))?.ToUpperInvariant()} editor {(string?) _.Element("HasEditorPermission")} metadata {string.Join(",", keys)}";
                })
                .Order(StringComparer.Ordinal));
    }

    static string Describe(byte[] binary)
    {
        try
        {
            var response = FssHttpBResponse.Read(binary);
            return string.Join(
                "; ",
                response.SubResponses.Select(_ =>
                {
                    if (_ is FailedSubResponse failed)
                    {
                        return $"{failed.Type} {failed.Error.Type} {failed.Error.Code}";
                    }

                    return $"{_.Type} ok";
                }));
        }
        catch (Exception exception) when (exception is FssHttpBException or InvalidDataException)
        {
            return $"unreadable: {exception.Message}";
        }
    }

    static List<string> Differences(Dictionary<string, string> expected, Dictionary<string, string> actual)
    {
        var differences = new List<string>();
        foreach (var key in expected.Keys.Union(actual.Keys))
        {
            var left = expected.GetValueOrDefault(key);
            var right = actual.GetValueOrDefault(key);
            if (left == right)
            {
                continue;
            }

            if (left != null &&
                right != null &&
                !key.StartsWith('@') &&
                key is not ("ErrorCode" or "ErrorMessage" or "FSSHTTPB" or "file" or "editors"))
            {
                // a list of keys: what one has and the other not
                var leftKeys = left.Split(',');
                var rightKeys = right.Split(',');
                differences.Add($"{key}: only SharePoint {string.Join(",", leftKeys.Except(rightKeys))}; only OfficeHost {string.Join(",", rightKeys.Except(leftKeys))}");
                continue;
            }

            differences.Add($"{key}: {left ?? "(none)"} / {right ?? "(none)"}");
        }

        return differences;
    }

    void Same(Exchange exchange, string label, string detail)
    {
        same++;
        lines.Add($"#{exchange.Number} {label}: same {detail}".TrimEnd());
    }

    void Differ(Exchange exchange, string label, List<string> differences)
    {
        different[label] = different.GetValueOrDefault(label) + 1;
        lines.Add($"#{exchange.Number} {label}: DIFFERENT (SharePoint / OfficeHost)");
        foreach (var difference in differences)
        {
            lines.Add($"    {difference}");
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
