/// <summary>
/// A difference of a SubResponse between SharePoint and OfficeHost: the facet (<see cref="ShadowComparer.Facets"/>) and
/// each server's value of it, or null when it had none.
/// </summary>
record ShadowDifference(string Facet, string? SharePoint, string? OfficeHost)
{
    public override string ToString()
    {
        if (SharePoint != null &&
            OfficeHost != null &&
            !Facet.StartsWith('@') &&
            Facet is not ("ErrorCode" or "ErrorMessage" or "FSSHTTPB" or "file" or "content" or "editors" or "HTTP" or "IsOnlyClientResult"))
        {
            // a list of keys: what one has and the other not
            var sharePoint = SharePoint.Split(',');
            var officeHost = OfficeHost.Split(',');
            return $"{Facet}: only SharePoint {string.Join(",", sharePoint.Except(officeHost))}; only OfficeHost {string.Join(",", officeHost.Except(sharePoint))}";
        }

        return $"{Facet}: {SharePoint ?? "(none)"} / {OfficeHost ?? "(none)"}";
    }
}

/// <summary>
/// What the shadow found of a sub-response, or of a whole exchange (a label of <c>cell storage</c> or
/// <c>IsOnlyClient</c>).
/// </summary>
/// <param name="Detail">The error code of a sub-response the same on both.</param>
/// <param name="Reason">Why it is a known difference or was not compared.</param>
/// <param name="ContentCompared">Whether the files of both copies of a partition were compared after it.</param>
/// <param name="Files">The two files of a content query whose files differ.</param>
record ShadowFinding(
    Exchange Exchange,
    string? Document,
    string Label,
    ShadowOutcome Outcome,
    IReadOnlyList<ShadowDifference> Differences,
    string? Detail = null,
    string? Reason = null,
    bool ContentCompared = false,
    (byte[] SharePoint, byte[] OfficeHost)? Files = null);

/// <summary>
/// Compares the answers of SharePoint and OfficeHost (docs/replay.md, "What it compares"): what is compared of a
/// SubResponse (its facets), the differences, the label of a sub-request, and the outcome of the differences.
/// </summary>
static class ShadowComparer
{
    public static XNamespace Soap { get; } = SoapBody.Soap;

    // the attributes whose values each server makes its own: compared by whether they are there
    static HashSet<string> volatileAttributes =
    [
        "TransitionID", "ResourceID", "Etag", "CreateTime", "LastModifiedTime", "ModifiedBy", "ServerTime", "UserName",
        "UserLogin", "UserEmailAddress", "UserSIPAddress", "HealthScore", "Url", "WebUrl"
    ];

    // the Editors Table Zip Stream Header ([MS-FSSHTTPB] 2.2.3.1.2.1.1)
    static byte[] editorsTableHeader = [0x1A, 0x5A, 0x3A, 0x30, 0x00, 0x00, 0x00, 0x00];

    /// <summary>
    /// The service a recorded request is for: <c>cellstorage</c>, <c>sharedaccess</c>, or null for the others.
    /// </summary>
    public static string? Kind(HttpMessage request)
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
    /// The documents the Requests of a cell storage request name, each with whether its Request makes the document: a
    /// Save As, whose put has the whole file and expects nothing.
    /// </summary>
    public static List<(ShadowDocument Document, bool Makes)> Documents(SoapBody request)
    {
        var documents = new List<(ShadowDocument, bool)>();
        foreach (var item in request.Document.Descendants(Soap + "Request"))
        {
            if (ShadowDocument.From((string?) item.Attribute("Url")) is { } document)
            {
                documents.Add((document, item.Elements(Soap + "SubRequest").Any(_ => Makes(_, request))));
            }
        }

        return documents;
    }

    /// <summary>
    /// Whether a sub-request is a Cell that puts a whole file expecting nothing, as a Save As does.
    /// </summary>
    static bool Makes(XElement subRequest, SoapBody body)
    {
        if ((string?) subRequest.Attribute("Type") != "Cell")
        {
            return false;
        }

        return RequestOf(body.Binary(subRequest.Element(Soap + "SubRequestData")))?.SubRequests
            .OfType<PutChangesRequest>()
            .Any(_ => _.ExpectedStorageIndex.IsNull) == true;
    }

    /// <summary>
    /// Whether a cell storage request starts a session on a document: a join of coauthoring, or a lock.
    /// </summary>
    public static bool Opens(SoapBody request) =>
        request.Document.Descendants(Soap + "SubRequestData")
            .Any(_ => (string?) _.Attribute("CoauthRequestType") == "JoinCoauthoring" ||
                      (string?) _.Attribute("SchemaLockRequestType") == "GetLock" ||
                      (string?) _.Attribute("ExclusiveLockRequestType") == "GetLock");

    /// <summary>
    /// The files that the queries of the content (Cells without a PartitionID) of a recorded exchange for a document
    /// answer in full: the document as it was, exactly.
    /// </summary>
    public static IEnumerable<byte[]> Contents(Exchange exchange, string key)
    {
        if (SoapBody.TryRead(exchange.Request) is not { } request ||
            SoapBody.TryRead(exchange.Response) is not { } response)
        {
            yield break;
        }

        var requests = request.Document.Descendants(Soap + "Request").ToList();
        foreach (var item in response.Document.Descendants(Soap + "Response"))
        {
            var requestToken = (string?) item.Attribute("RequestToken");
            var matching = requests.FirstOrDefault(_ => (string?) _.Attribute("RequestToken") == requestToken);
            if (matching == null ||
                ShadowDocument.From((string?) matching.Attribute("Url"))?.Key != key)
            {
                continue;
            }

            var contentTokens = matching.Elements(Soap + "SubRequest")
                .Where(_ => (string?) _.Attribute("Type") == "Cell" &&
                            PartitionOf(_.Element(Soap + "SubRequestData")) == FssHttpCells.ContentPartition)
                .Select(_ => (string?) _.Attribute("SubRequestToken"))
                .ToHashSet();
            foreach (var subResponse in item.Elements(Soap + "SubResponse").Where(_ => contentTokens.Contains((string?) _.Attribute("SubRequestToken"))))
            {
                if (FileOf(response.Binary(subResponse.Element(Soap + "SubResponseData"))) is { } file)
                {
                    yield return file;
                }
            }
        }
    }

    /// <summary>
    /// The version of the file in the Etags of a recorded answer (<c>"{guid},N"</c>): the highest N, or null.
    /// </summary>
    public static long? EtagVersion(Exchange exchange)
    {
        if (SoapBody.TryRead(exchange.Response) is not { } response)
        {
            return null;
        }

        long? version = null;
        foreach (var etag in response.Document.Descendants(Soap + "SubResponseData").Select(_ => (string?) _.Attribute("Etag")))
        {
            if (etag != null &&
                long.TryParse(etag.Trim('"').Split(',')[^1], NumberStyles.None, CultureInfo.InvariantCulture, out var value) &&
                (version == null || value > version))
            {
                version = value;
            }
        }

        return version;
    }

    /// <summary>
    /// The file a Cell answer holds in full, or null.
    /// </summary>
    public static byte[]? FileOf(byte[]? binary)
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

    public static Guid PartitionOf(XElement? data)
    {
        if (Guid.TryParse((string?) data?.Attribute("PartitionID"), out var partition))
        {
            return partition;
        }

        return FssHttpCells.ContentPartition;
    }

    public static FssHttpBRequest? RequestOf(byte[]? binary)
    {
        if (binary == null)
        {
            return null;
        }

        try
        {
            return FssHttpBRequest.Read(binary);
        }
        catch (Exception exception) when (exception is FssHttpBException or InvalidDataException)
        {
            return null;
        }
    }

    public static FssHttpBResponse? ResponseOf(byte[]? binary)
    {
        if (binary == null)
        {
            return null;
        }

        try
        {
            return FssHttpBResponse.Read(binary);
        }
        catch (Exception exception) when (exception is FssHttpBException or InvalidDataException)
        {
            return null;
        }
    }

    /// <summary>
    /// The result of an IsOnlyClient answer.
    /// </summary>
    public static string? IsOnlyClientResult(HttpMessage response)
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

    /// <summary>
    /// The sub-request type, with its request type (<c>Coauth JoinCoauthoring</c>), or the FSSHTTPB sub-requests of a
    /// Cell (<c>Cell QueryChanges</c>), with whether the ids of SharePoint in it were translated
    /// (<see cref="ReplayCells"/>).
    /// </summary>
    public static string Label(XElement? subRequest, SoapBody request, string? status)
    {
        if (subRequest == null)
        {
            return "?";
        }

        var type = (string?) subRequest.Attribute("Type") ?? "?";
        var data = subRequest.Element(Soap + "SubRequestData");
        if (type == "Cell")
        {
            if (RequestOf(request.Binary(data)) is not { } fssHttpB)
            {
                return "Cell (unreadable)";
            }

            var types = string.Join(" ", fssHttpB.SubRequests.Select(_ => _.Type));
            if (status != null)
            {
                return $"Cell {types} ({status})";
            }

            return $"Cell {types}";
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
    /// query answers in full, or the editors of an editors table.
    /// </summary>
    public static Dictionary<string, string> Facets(XElement subResponse, SoapBody body)
    {
        var facets = new Dictionary<string, string>
        {
            ["ErrorCode"] = (string?) subResponse.Attribute("ErrorCode") ?? ""
        };
        if ((string?) subResponse.Attribute("ErrorMessage") is { } message)
        {
            facets["ErrorMessage"] = message;
        }

        var data = subResponse.Element(Soap + "SubResponseData");
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
                    facets["file"] = Hash(file);
                }
            }
        }

        return facets;
    }

    public static string Hash(byte[] file) => Convert.ToHexString(SHA256.HashData(file))[..12];

    /// <summary>
    /// An editors table by what the servers agree on: each editor's CacheID, permission and metadata keys. Not its login
    /// name (the users are mapped), its timeout (OfficeHost keeps clients an hour at least), nor the layout of its XML.
    /// </summary>
    public static string? EditorsOf(byte[] file)
    {
        if (!file.AsSpan().StartsWith(editorsTableHeader))
        {
            return null;
        }

        XDocument table;
        try
        {
            using var deflate = new DeflateStream(new MemoryStream(file, editorsTableHeader.Length, file.Length - editorsTableHeader.Length), CompressionMode.Decompress);
            table = XDocument.Load(deflate);
        }
        catch (Exception exception) when (exception is InvalidDataException or XmlException)
        {
            return null;
        }

        // another stream with the same header, as Word's metadata
        if (table.Root?.Name.LocalName != "EditorsTable")
        {
            return null;
        }

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

    public static List<ShadowDifference> Differences(Dictionary<string, string> sharePoint, Dictionary<string, string> officeHost) =>
        sharePoint.Keys.Union(officeHost.Keys)
            .Select(_ => new ShadowDifference(_, sharePoint.GetValueOrDefault(_), officeHost.GetValueOrDefault(_)))
            .Where(_ => _.SharePoint != _.OfficeHost)
            .ToList();

    /// <summary>
    /// The outcome of the differences of a sub-response: the same; not compared, for a Cell of a client that knew the
    /// document from Office's cache; a known difference, when the rules expect them all; OfficeHost failed, when it
    /// failed a sub-request SharePoint did not; or different.
    /// </summary>
    public static ShadowOutcome Outcome(string label, IReadOnlyList<ShadowDifference> differences, IReadOnlyList<ShadowRule> rules, out string? reason)
    {
        reason = null;
        if (differences.Count == 0)
        {
            return ShadowOutcome.Same;
        }

        if (label.EndsWith($"({ReplayCells.FromCache})", StringComparison.Ordinal))
        {
            reason = ReplayCells.FromCache;
            return ShadowOutcome.NotCompared;
        }

        var matched = differences
            .Select(difference => rules.FirstOrDefault(_ => _.Matches(label, difference.Facet, difference.SharePoint, difference.OfficeHost)))
            .ToList();
        if (matched.All(_ => _ != null))
        {
            reason = string.Join("; ", matched.Select(_ => _!.Reason).Distinct());
            return ShadowOutcome.KnownDifference;
        }

        if (differences.FirstOrDefault(_ => _.Facet == "ErrorCode") is { SharePoint: "Success", OfficeHost: { } failed } &&
            failed != "Success")
        {
            reason = $"OfficeHost answered {failed} where SharePoint succeeded";
            return ShadowOutcome.OfficeHostFailed;
        }

        return ShadowOutcome.Different;
    }
}
