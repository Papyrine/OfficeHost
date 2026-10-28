using System.IO.Compression;

/// <summary>
/// The file and user a cell storage request is for, from the token in its path, and the URL of its site.
/// </summary>
/// <param name="Token">The token in the path of the site, and when it expires.</param>
record FssHttpSite(OfficeAccess Access, string WebUrl, string Token, DateTimeOffset Expires)
{
    public string FileId => Access.FileId;

    /// <summary>
    /// The site of a Save As that replaces a file the site cannot reach otherwise: its content is written, not read.
    /// </summary>
    public bool Replacing { get; init; }

    /// <summary>
    /// The site as it reaches another file, one it made by Save As: the same user, with the same rights.
    /// </summary>
    public FssHttpSite For(string fileId) =>
        this with
        {
            Access = Access with
            {
                FileId = fileId
            }
        };

    /// <summary>
    /// The ResourceID and TransitionID of the file ([MS-FSSHTTP] 2.3.1.7): a GUID of its id.
    /// </summary>
    public Guid ResourceId => ResourceIdOf(FileId);

    public static Guid ResourceIdOf(string fileId) => new(SHA256.HashData(Encoding.UTF8.GetBytes(fileId)).AsSpan(0, 16));
}

/// <summary>
/// The cell storage service ([MS-FSSHTTP] 3.1.4): runs the sub-requests of each Request, in order, as their dependencies
/// allow (2.2.5.3), and answers each with a SubResponse.
/// </summary>
class CellStorageService(
    IOfficeStorage storage,
    FssHttpCells cells,
    CoauthSessions sessions,
    FssHttpTargets targets,
    SaveAsFiles saveAs,
    IFssHttpStore store,
    IOptions<OfficeHostOptions> options,
    ILogger<CellStorageService> log)
{
    static XNamespace soap = SoapMessage.Soap;
    // E_FAIL, the HResult of SharePoint with an error code (MS-FSSHTTP 4.1.2)
    const long failHResult = 2147500037;
    const string success = "Success";

    const int maxRequests = 16;

    const int maxSubRequests = 128;

    /// <summary>
    /// The most clients sharing the lock, and entries in the editors table: SharePoint lets 99 users coauthor, the
    /// editors table also has the readers.
    /// </summary>
    const int maxClients = 99;

    const int maxEditors = 256;

    // the Editors Table Zip Stream Header ([MS-FSSHTTPB] 2.2.3.1.2.1.1)
    static byte[] editorsTableHeader = [0x1A, 0x5A, 0x3A, 0x30, 0x00, 0x00, 0x00, 0x00];

    public async Task<XElement> Execute(SoapMessage message, FssHttpSite site, MtomResponse mtom, Cancel cancel)
    {
        var body = message.Document.Root?.Element(SoapMessage.Envelope + "Body") ??
                   throw new InvalidDataException("No SOAP body");
        // Office sends a few requests of a few sub-requests: many more are refused before any runs
        var requests = body.Descendants(soap + "Request").ToList();
        if (requests.Count > maxRequests ||
            requests.Any(_ => _.Elements(soap + "SubRequest").Skip(maxSubRequests).Any()))
        {
            throw new InvalidDataException($"More than {maxRequests} requests, or {maxSubRequests} sub-requests in a request");
        }

        var responses = new List<XElement>();
        foreach (var request in requests)
        {
            responses.Add(await Target(message, request, site, mtom, cancel));
        }

        return new(
            SoapMessage.Envelope + "Envelope",
            new XAttribute(XNamespace.Xmlns + "s", SoapMessage.Envelope),
            new XElement(
                SoapMessage.Envelope + "Body",
                // 2.2.3.7: SharePoint answers MinorVersion 3, which XORs the signatures of small zip items (MS-FSSHTTPD 2.4.1)
                new XElement(soap + "ResponseVersion", new XAttribute("Version", "2"), new XAttribute("MinorVersion", "3")),
                new XElement(soap + "ResponseCollection", new XAttribute("WebUrl", site.WebUrl), responses)));
    }

    /// <summary>
    /// A Request for the file its Url names (<see cref="FssHttpTargets"/>). A Save As is a Request to the URL of a new
    /// name in the site, whose Cell puts the whole file (docs/fsshttp-notes.md): the file is created, or an unlocked file
    /// of that name taken over, and the Request runs on it; a file it created is deleted again when the Cell fails. A
    /// name that names no file fails every sub-request, as on SharePoint.
    /// </summary>
    async Task<XElement> Target(SoapMessage message, XElement request, FssHttpSite site, MtomResponse mtom, Cancel cancel)
    {
        var target = await targets.Find(site, FssHttpTargets.NameOf(site, (string?) request.Attribute("Url")), cancel);
        if (target.Site != null)
        {
            return await Request(message, request, target.Site, mtom, cancel);
        }

        var putTokens = request.Elements(soap + "SubRequest")
            .Where(_ => (string?) _.Attribute("Type") == "Cell" && Puts(message, _))
            .Select(_ => (string?) _.Attribute("SubRequestToken") ?? "")
            .ToList();
        if (target.SaveAsName == null ||
            putTokens.Count == 0)
        {
            return Failed(request, "FileNotExistsOrCannotBeCreated");
        }

        var file = target.SaveAsExisting;
        var created = false;
        if (file == null)
        {
            file = await storage.CreateSibling(site.FileId, target.SaveAsName, Stream.Null, false, cancel);
            created = file != null;
        }
        else if ((await sessions.Get(file.Id, cancel)).Lock() != null ||
                 await sessions.ForeignLock(file.Id, cancel) != null)
        {
            return Failed(request, "FileAlreadyLockedOnServer");
        }

        if (file == null)
        {
            // another file took the name meanwhile
            return Failed(request, "FileAlreadyExists");
        }

        var fileSite = site.For(file.Id);
        if (!created)
        {
            // the file replaced is not read through the site: its cell is written again from the file with new ids, so
            // that a Put cannot refer to its data elements, and its content is not queried until the site made it
            await cells.Reset(file.Id, FssHttpCells.ContentPartition, cancel);
            fileSite = fileSite with
            {
                Replacing = true
            };
        }

        var response = await Request(message, request, fileSite, mtom, cancel);
        var failed = response.Elements(soap + "SubResponse")
            .Any(_ => putTokens.Contains((string?) _.Attribute("SubRequestToken") ?? "") &&
                      (string?) _.Attribute("ErrorCode") != success);
        if (failed)
        {
            if (created)
            {
                await storage.Delete(file.Id, Cancel.None);
                await store.Delete(file.Id, Cancel.None);
            }

            return response;
        }

        // the site reaches the file once it wrote it
        await saveAs.Add(site, file.Id, cancel);
        return response;
    }

    /// <summary>
    /// Whether a Cell sub-request puts changes.
    /// </summary>
    static bool Puts(SoapMessage message, XElement subRequest)
    {
        if (subRequest.Element(soap + "SubRequestData") is not { } data)
        {
            return false;
        }

        try
        {
            return message.FssHttpB(data)?.SubRequests.Any(_ => _ is PutChangesRequest) == true;
        }
        catch (Exception exception) when (exception is FssHttpBException or InvalidDataException)
        {
            return false;
        }
    }

    /// <summary>
    /// A Response whose sub-requests all fail with the error code.
    /// </summary>
    static XElement Failed(XElement request, string errorCode) =>
        new(
            soap + "Response",
            new XAttribute("Url", (string?) request.Attribute("Url") ?? ""),
            new XAttribute("RequestToken", (string?) request.Attribute("RequestToken") ?? "1"),
            new XAttribute("HealthScore", "0"),
            request.Elements(soap + "SubRequest")
                .Select(_ => new XElement(
                    soap + "SubResponse",
                    new XAttribute("SubRequestToken", (string?) _.Attribute("SubRequestToken") ?? ""),
                    new XAttribute("ErrorCode", errorCode),
                    new XAttribute("HResult", failHResult),
                    new XElement(soap + "SubResponseData"))));

    async Task<XElement> Request(SoapMessage message, XElement request, FssHttpSite site, MtomResponse mtom, Cancel cancel)
    {
        var results = new Dictionary<string, string>();
        var subResponses = new List<XElement>();
        foreach (var subRequest in request.Elements(soap + "SubRequest"))
        {
            var token = (string?) subRequest.Attribute("SubRequestToken") ?? "";
            var data = new XElement(soap + "SubResponseData");
            var errorCode = Dependency(subRequest, results) ??
                            await SubRequest(message, subRequest, site, data, mtom, cancel);
            results[token] = errorCode;
            long hResult = 0;
            if (errorCode != success)
            {
                hResult = failHResult;
            }

            var subResponse = new XElement(
                soap + "SubResponse",
                new XAttribute("SubRequestToken", token),
                new XAttribute("ErrorCode", errorCode));
            // SharePoint names the user who has the lock in the ErrorMessage of a lock refused, which Excel shows in its
            // File In Use; not in that of a rename refused
            if (errorCode == "FileAlreadyLockedOnServer" &&
                (string?) subRequest.Attribute("Type") is "Coauth" or "SchemaLock" or "ExclusiveLock" &&
                LockHolder(await sessions.Get(site.FileId, cancel)) is { } holder)
            {
                subResponse.Add(new XAttribute("ErrorMessage", holder));
            }

            subResponse.Add(new XAttribute("HResult", hResult), data);
            subResponses.Add(subResponse);
        }

        return new(
            soap + "Response",
            new XAttribute("Url", (string?) request.Attribute("Url") ?? ""),
            new XAttribute("RequestToken", (string?) request.Attribute("RequestToken") ?? "1"),
            new XAttribute("HealthScore", "0"),
            new XAttribute("ResourceID", site.ResourceId),
            subResponses);
    }

    /// <summary>
    /// The error code of a sub-request its dependency keeps from running (2.2.5.2, 2.2.5.3), or null to run it. OnExecute
    /// runs it after its dependency ran, failed or not, as SharePoint Server SE does: not after one kept from running by
    /// a failed OnSuccess or OnFail, but after one kept from running by OnNotSupported, as in Word's open
    /// (docs/fsshttp-notes.md).
    /// </summary>
    static string? Dependency(XElement subRequest, Dictionary<string, string> results)
    {
        var dependsOn = (string?) subRequest.Attribute("DependsOn");
        if (dependsOn == null ||
            !results.TryGetValue(dependsOn, out var dependency))
        {
            return null;
        }

        return (string?) subRequest.Attribute("DependencyType") switch
        {
            "OnExecute" when dependency is "DependentOnlyOnSuccessRequestFailed" or "DependentOnlyOnFailRequestSucceeded" or "DependentRequestNotExecuted" => "DependentRequestNotExecuted",
            null or "OnExecute" => null,
            "OnSuccess" when dependency != success => "DependentOnlyOnSuccessRequestFailed",
            "OnFail" when dependency == success => "DependentOnlyOnFailRequestSucceeded",
            "OnNotSupported" when dependency != "RequestNotSupported" => "DependentOnlyOnNotSupportedRequestGetSupported",
            "OnSuccessOrNotSupported" when dependency is not (success or "RequestNotSupported") => "DependentOnlyOnSuccessRequestFailed",
            "OnSuccess" or "OnFail" or "OnNotSupported" or "OnSuccessOrNotSupported" => null,
            _ => "InvalidRequestDependencyType"
        };
    }

    async Task<string> SubRequest(SoapMessage message, XElement subRequest, FssHttpSite site, XElement data, MtomResponse mtom, Cancel cancel)
    {
        var requestData = subRequest.Element(soap + "SubRequestData");
        try
        {
            return (string?) subRequest.Attribute("Type") switch
            {
                "Cell" => await Cell(message, requestData, site, data, mtom, cancel),
                "Coauth" => await Coauth(requestData, site, data, cancel),
                "SchemaLock" => await SchemaLock(requestData, site, data, cancel),
                "ExclusiveLock" => await ExclusiveLock(requestData, site, data, cancel),
                "EditorsTable" => await EditorsTable(requestData, site, cancel),
                "WhoAmI" => WhoAmI(site, data),
                "ServerTime" => ServerTime(data),
                "GetDocMetaInfo" => await GetDocMetaInfo(site, data, cancel),
                "AmIAlone" => await AmIAlone(site, data, cancel),
                "LockStatus" => await LockStatus(site, data, cancel),
                "FileOperation" => await FileOperation(requestData, site, cancel),
                "Properties" => Properties(requestData, data),
                _ => "RequestNotSupported"
            };
        }
        catch (Exception exception) when (exception is InvalidDataException or FssHttpBException or FormatException)
        {
            return "InvalidArgument";
        }
    }

    /// <summary>
    /// A Cell sub-request (3.1.4.2): the FSSHTTPB of its data, whose sub-requests target the partition of the
    /// PartitionID attribute (2.3.3.1) or their own Target Partition Id ([MS-FSSHTTPB] 2.2.2.1.1). The FSSHTTPB answer
    /// is an MTOM part; a Put Changes that fails (a coherency failure) fails in it, with the error code CellRequestFail,
    /// as on SharePoint Server SE.
    /// </summary>
    async Task<string> Cell(SoapMessage message, XElement? requestData, FssHttpSite site, XElement data, MtomResponse mtom, Cancel cancel)
    {
        if (requestData == null ||
            message.FssHttpB(requestData) is not { } request)
        {
            return success;
        }

        var file = await storage.GetFile(site.FileId, cancel);
        if (file == null)
        {
            return "CellRequestFail";
        }

        // an Etag of another version of the file fails, as on SharePoint Server SE (Office sends none)
        if (Attribute(requestData, "Etag") is { } etag &&
            etag != Etag(site, file))
        {
            return "CellRequestFail";
        }

        var partitionAttribute = (string?) requestData.Attribute("PartitionID");
        var defaultPartition = Guid.Empty;
        if (partitionAttribute != null)
        {
            defaultPartition = Guid.Parse(partitionAttribute);
        }

        // the Cell of a Save As takes the exclusive lock it names, which its ExclusiveLock ConvertToSchemaJoinCoauth then
        // turns into a coauthoring session, as on SharePoint Server SE
        var tookExclusiveLock = false;
        if (Attribute(requestData, "ExclusiveLockID") is { } exclusiveLockId &&
            Int(requestData, "Timeout") is { } timeout &&
            request.SubRequests.Any(_ => _ is PutChangesRequest))
        {
            var locked = await TakeExclusiveLock(site, exclusiveLockId, timeout, cancel);
            if (locked != success)
            {
                return locked;
            }

            tookExclusiveLock = true;
        }

        var subResponses = new List<SubResponse>();
        var elements = new Dictionary<ExtendedGuid, DataElement>();
        var package = request.DataElements?.Elements ?? [];
        foreach (var subRequest in request.SubRequests)
        {
            var partition = subRequest.TargetPartition ?? defaultPartition;
            // the partitions OfficeHost keeps its own state in are not cells
            if (partition == CoauthSessions.SessionPartition ||
                partition == SaveAsFiles.Partition)
            {
                subResponses.Add(new FailedSubResponse(subRequest.RequestId, subRequest.Type, ResponseError.Cell(CellErrorCode.InvalidPartition)));
                continue;
            }

            switch (subRequest)
            {
                case QueryAccessRequest access:
                    var write = ResponseError.Success;
                    if (!site.Access.CanWrite)
                    {
                        // E_ACCESSDENIED
                        write = new(ResponseErrorType.HResult, 0x80070005);
                    }

                    subResponses.Add(new QueryAccessResponse(access.RequestId, ResponseError.Success, write));
                    break;
                case QueryChangesRequest when site.Replacing && partition == FssHttpCells.ContentPartition:
                    subResponses.Add(new FailedSubResponse(subRequest.RequestId, subRequest.Type, ResponseError.Cell(CellErrorCode.StorageWriteOnly)));
                    break;
                case QueryChangesRequest query:
                    var cell = await Partition(site, partition, cancel);
                    var (response, sent) = CellSync.Query(query, cell.State.Storage!);
                    subResponses.Add(response);
                    foreach (var element in sent)
                    {
                        elements.TryAdd(element.Id, element);
                    }

                    break;
                case PutChangesRequest put:
                    var (errorCode, putResponse) = await Put(put, package, partition, requestData, site, cancel);
                    if (errorCode != null)
                    {
                        return errorCode;
                    }

                    subResponses.Add(putResponse!);
                    break;
                case AllocateExtendedGuidRangeRequest allocate:
                    // [MS-FSSHTTPB] 2.2.3.1.4: the range ends between 1,000 and 100,000
                    subResponses.Add(new AllocateExtendedGuidRangeResponse(allocate.RequestId, Guid.NewGuid(), 1, Math.Clamp(allocate.Count, 999UL, 99_999UL) + 1));
                    break;
            }
        }

        var fssHttpB = new FssHttpBResponse(subResponses, new(elements.Values.ToList()), ProtocolVersion: 13).ToMemory();
        // the properties of the file, for a Cell of the content only, as SharePoint Server SE answers
        // (docs/fsshttp-notes.md, "Word's session replayed")
        if (requestData.Attribute("GetFileProps")?.Value is "true" or "True" &&
            defaultPartition == FssHttpCells.ContentPartition)
        {
            file = await storage.GetFile(site.FileId, cancel) ?? file;
            data.Add(
                new XAttribute("Etag", Etag(site, file)),
                new XAttribute("CreateTime", file.LastModified.ToFileTime()),
                new XAttribute("LastModifiedTime", file.LastModified.ToFileTime()),
                new XAttribute("ModifiedBy", site.Access.UserName ?? site.Access.UserId));
            // with a query that sends changes, not with one that has none, nor with a put
            if (elements.Count > 0 &&
                !request.SubRequests.Any(_ => _ is PutChangesRequest))
            {
                data.Add(new XAttribute("HaveOnlyDemotionChanges", "False"));
            }
        }

        // the Cell of a Save As holds the exclusive lock it took
        if (tookExclusiveLock)
        {
            data.Add(new XAttribute("LockType", "ExclusiveLock"));
        }

        data.Add(mtom.Include(fssHttpB));
        if (subResponses.Any(_ => _ is FailedSubResponse))
        {
            return "CellRequestFail";
        }

        return success;
    }

    static string Etag(FssHttpSite site, OfficeFile file) => $"\"{{{site.ResourceId.ToString().ToUpperInvariant()}}},{file.Version}\"";

    /// <summary>
    /// The cell of a partition: the content in step with the file, the editors table from the coauthoring session, and
    /// any other partition (the metadata) as Office last put it, or with nothing in it.
    /// </summary>
    async Task<StoredCell> Partition(FssHttpSite site, Guid partition, Cancel cancel)
    {
        if (partition == FssHttpCells.ContentPartition)
        {
            return await cells.GetContent(site.FileId, cancel) ??
                   throw new InvalidDataException($"No file {site.FileId}");
        }

        if (partition == FssHttpCells.EditorsTablePartition)
        {
            return await EditorsTableCell(site.FileId, cancel);
        }

        for (var attempt = 0; ; attempt++)
        {
            var cell = await cells.Get(site.FileId, partition, cancel);
            if (cell.State.Storage != null)
            {
                return cell;
            }

            var ids = cell.State.Ids();
            var empty = CellSync.Empty(ids);
            var state = new CellState(empty, ids.Guid, ids.Next, null);
            var stored = await cells.TryUpdate(site.FileId, partition, cell.Version, state, cancel);
            if (stored != null || attempt >= 3)
            {
                return stored ?? new StoredCell(null, state);
            }
        }
    }

    /// <summary>
    /// A Put Changes: the error code of the Cell sub-request when the lock or the access forbids it, otherwise the
    /// sub-response. The content is stored first without a file version, then written to the file, then stored with
    /// the version of the file: a failure in between leaves a cell newer than the file, which
    /// <see cref="FssHttpCells.GetContent"/> writes to the file.
    /// </summary>
    async Task<(string? ErrorCode, SubResponse? Response)> Put(PutChangesRequest put, IReadOnlyList<DataElement> package, Guid partition, XElement requestData, FssHttpSite site, Cancel cancel)
    {
        if (!site.Access.CanWrite)
        {
            return ("FileUnauthorizedAccess", null);
        }

        if (partition == FssHttpCells.EditorsTablePartition)
        {
            return (null, new FailedSubResponse(put.RequestId, RequestType.PutChanges, ResponseError.Cell(CellErrorCode.StorageReadOnly)));
        }

        var lockId = (string?) requestData.Attribute("BypassLockID") ?? (string?) requestData.Attribute("SchemaLockID");
        if (partition == FssHttpCells.ContentPartition &&
            !await sessions.CanWrite(site.FileId, lockId, site.Access.UserId, cancel))
        {
            return ("FileAlreadyLockedOnServer", null);
        }

        for (var attempt = 0; attempt < 8; attempt++)
        {
            StoredCell cell;
            if (partition == FssHttpCells.ContentPartition)
            {
                cell = await cells.GetContent(site.FileId, cancel) ??
                       throw new InvalidDataException($"No file {site.FileId}");
            }
            else
            {
                cell = await cells.Get(site.FileId, partition, cancel);
            }

            var outcome = CellSync.Put(put, package, cell.State);
            if (outcome.State == null)
            {
                return (null, outcome.Response);
            }

            if (partition != FssHttpCells.ContentPartition)
            {
                if (await cells.TryUpdate(site.FileId, partition, cell.Version, outcome.State, cancel) != null)
                {
                    return (null, outcome.Response);
                }

                continue;
            }

            byte[] bytes;
            try
            {
                bytes = FileCell.Read(outcome.State.Storage!).Root.ToArray();
            }
            catch (FssHttpDException)
            {
                return (null, new FailedSubResponse(put.RequestId, RequestType.PutChanges, ResponseError.Cell(CellErrorCode.InvalidObject)));
            }

            var pending = await cells.TryUpdate(site.FileId, partition, cell.Version, outcome.State with { FileVersion = null, BaseFileVersion = cell.State.FileVersion }, cancel);
            if (pending == null)
            {
                continue;
            }

            var written = await cells.WriteFile(site.FileId, pending, bytes, cancel);
            if (options.Value.OnFileSaved != null &&
                await storage.GetFile(site.FileId, cancel) is { } file &&
                file.Version == written.State.FileVersion)
            {
                await WopiHandler.NotifySaved(options.Value, log, new(file, site.Access, [site.Access.UserId], false, false));
            }

            return (null, outcome.Response);
        }

        return (null, new FailedSubResponse(put.RequestId, RequestType.PutChanges, ResponseError.Cell(CellErrorCode.CoherencyFailure)));
    }

    static string? Attribute(XElement? element, string name) => (string?) element?.Attribute(name);

    static int? Int(XElement? element, string name)
    {
        var value = Attribute(element, name);
        if (value == null)
        {
            return null;
        }

        // a value out of range is an invalid argument, as one that is not a number
        if (!int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var result))
        {
            throw new FormatException($"{name} is not a number of the range of an int: {value}");
        }

        return result;
    }

    static bool Bool(XElement? element, string name) => Attribute(element, name) is "true" or "True" or "1";

    static string Status(Session session)
    {
        if (session.Clients.Count > 1)
        {
            return "Coauthoring";
        }

        return "Alone";
    }

    /// <summary>
    /// Joins the shared lock of the schema lock id (3.1.4.3.1, 3.1.4.4.1): the client is added, or refreshed when it is
    /// there, and gets an entry in the editors table. A second client starts a coauthoring transition.
    /// </summary>
    SessionResult Join(Session session, FssHttpSite site, string clientId, string schemaLockId, int timeout)
    {
        if (!site.Access.CanWrite)
        {
            return new("FileUnauthorizedAccess", session);
        }

        if (session.ExclusiveLockId != null ||
            (session.SchemaLockId != null && session.SchemaLockId != schemaLockId))
        {
            return new("FileAlreadyLockedOnServer", session);
        }

        var expires = sessions.Now.AddSeconds(timeout);
        var known = session.Client(clientId) != null;
        if (!known &&
            session.Clients.Count >= maxClients)
        {
            return new("NumberOfCoauthorsReachedMax", session);
        }

        var clients = session.Clients.Where(_ => _.ClientId != clientId).Append(new(clientId, expires, timeout)).ToList();
        var editor = session.Editors.FirstOrDefault(_ => _.ClientId == clientId);
        var editors = session.Editors.Where(_ => _.ClientId != clientId)
            .Append(new(clientId, site.Access.UserId, site.Access.UserName, true, expires, editor?.Metadata ?? JoinMetadata()))
            .ToList();
        var pending = session.TransitionPending || (!known && session.Clients.Count > 0);
        return new(success, session with
        {
            SchemaLockId = schemaLockId,
            Clients = clients,
            TransitionPending = pending,
            Editors = editors
        });
    }

    /// <summary>
    /// The client leaves the shared lock (3.1.4.3.2 Exit Coauthoring Session, 3.1.4.4.2 Release Lock). An exit is a
    /// Success whenever the client is not in the session, even under another schema lock id, which it then leaves as it
    /// is, as SharePoint Server SE answers (docs/fsshttp-notes.md).
    /// </summary>
    SessionResult Leave(Session session, string clientId, string schemaLockId, bool coauth)
    {
        if (coauth &&
            session.ExclusiveLockId == null &&
            (session.SchemaLockId != schemaLockId || session.Client(clientId) == null))
        {
            return new(success, session);
        }

        if (session.ExclusiveLockId != null ||
            (session.SchemaLockId != null && session.SchemaLockId != schemaLockId))
        {
            return new("FileAlreadyLockedOnServer", session);
        }

        if (session.Client(clientId) == null)
        {
            // 3.1.4.4.2: InvalidCoauthSession with others in the session, Success without
            if (session.Clients.Count > 0)
            {
                return new("InvalidCoauthSession", session);
            }

            return new(success, session);
        }

        return new(success, (session with
        {
            Clients = session.Clients.Where(_ => _.ClientId != clientId).ToList(),
            Editors = session.Editors.Where(_ => _.ClientId != clientId).ToList()
        }).Purge(sessions.Now));
    }

    SessionResult ConvertToExclusive(Session session, FssHttpSite site, string clientId, string schemaLockId, string? exclusiveLockId, int timeout, bool release)
    {
        if (session.ExclusiveLockId != null ||
            session.SchemaLockId != schemaLockId ||
            session.Client(clientId) == null ||
            exclusiveLockId == null)
        {
            return new("InvalidCoauthSession", session);
        }

        if (session.Clients.Count > 1)
        {
            if (release)
            {
                // the client leaves, but the session keeps the others: the error code is that of a success to store
                return new("ExitCoauthSessionAsConvertToExclusiveFailed", session);
            }

            return new("MultipleClientsInCoauthSession", session);
        }

        return new(success, Session.None with
        {
            ExclusiveLockId = exclusiveLockId,
            ExclusiveExpires = sessions.Now.AddSeconds(timeout),
            Editors = session.Editors,
            ExclusiveOwner = site.Access.UserName ?? site.Access.UserId,
            ExclusiveUserId = site.Access.UserId
        });
    }

    /// <summary>
    /// Whether the lock is available: not held by another session lock, nor by a WOPI client.
    /// </summary>
    async Task<string> Availability(FssHttpSite site, Session session, string? schemaLockId, string? exclusiveLockId, Cancel cancel)
    {
        if ((session.ExclusiveLockId != null && session.ExclusiveLockId != exclusiveLockId) ||
            (session.SchemaLockId != null && session.SchemaLockId != schemaLockId) ||
            await sessions.ForeignLock(site.FileId, cancel) != null)
        {
            return "FileAlreadyLockedOnServer";
        }

        return success;
    }

    /// <summary>
    /// Changes the session, and, with "ExitCoauthSessionAsConvertToExclusiveFailed", takes the client out of it too.
    /// </summary>
    async Task<SessionResult> Change(FssHttpSite site, Func<Session, SessionResult> change, string? clientId, Cancel cancel)
    {
        var result = await sessions.Change(site.FileId, change, cancel);
        if (result.ErrorCode != "ExitCoauthSessionAsConvertToExclusiveFailed" ||
            clientId == null)
        {
            return result;
        }

        await sessions.Change(
            site.FileId,
            _ => new(success, _ with
            {
                Clients = _.Clients.Where(_ => _.ClientId != clientId).ToList(),
                Editors = _.Editors.Where(_ => _.ClientId != clientId).ToList()
            }),
            cancel);
        return result;
    }

    async Task<string> Coauth(XElement? requestData, FssHttpSite site, XElement data, Cancel cancel)
    {
        var clientId = Attribute(requestData, "ClientID");
        var schemaLockId = Attribute(requestData, "SchemaLockID");
        if (clientId == null || schemaLockId == null)
        {
            return "InvalidArgument";
        }

        if (TimeoutTooLong(requestData))
        {
            return "LockRequestFail";
        }

        var timeout = CoauthSessions.Timeout(Int(requestData, "Timeout"));
        switch (Attribute(requestData, "CoauthRequestType"))
        {
            case "JoinCoauthoring":
            case "RefreshCoauthoring":
            {
                var result = await Change(site, _ => Join(_, site, clientId, schemaLockId, timeout), null, cancel);
                if (result.ErrorCode == success)
                {
                    data.Add(
                        new XAttribute("LockType", "SchemaLock"),
                        new XAttribute("CoauthStatus", Status(result.Session)));
                    // with a join only: SharePoint Server SE answers a refresh without it (docs/fsshttp-notes.md)
                    if (Attribute(requestData, "CoauthRequestType") == "JoinCoauthoring")
                    {
                        data.Add(new XAttribute("TransitionID", site.ResourceId));
                    }
                }

                return result.ErrorCode;
            }
            case "ExitCoauthoring":
                return (await Change(site, _ => Leave(_, clientId, schemaLockId, true), null, cancel)).ErrorCode;
            case "ConvertToExclusive":
                return (await Change(
                    site,
                    _ => ConvertToExclusive(_, site, clientId, schemaLockId, Attribute(requestData, "ExclusiveLockID"), timeout, Bool(requestData, "ReleaseLockOnConversionToExclusiveFailure")),
                    clientId,
                    cancel)).ErrorCode;
            case "CheckLockAvailability":
                return await Availability(site, await sessions.Get(site.FileId, cancel), schemaLockId, null, cancel);
            case "MarkTransitionComplete":
                return (await Change(
                    site,
                    _ =>
                    {
                        // a Success without the client too, as SharePoint Server SE answers
                        if (_.Client(clientId) == null)
                        {
                            return new(success, _);
                        }

                        return new(success, _ with
                        {
                            TransitionPending = false
                        });
                    },
                    null,
                    cancel)).ErrorCode;
            case "GetCoauthoringStatus":
            {
                // the poll of a coauthoring client, every 8 seconds: it keeps the client in the session
                var result = await Change(
                    site,
                    _ =>
                    {
                        if (_.Client(clientId) is not { } client)
                        {
                            return new(success, _);
                        }

                        var expires = sessions.Now.AddSeconds(client.TimeoutSeconds);
                        return new(success, _ with
                        {
                            Clients =
                            [
                                .. _.Clients
                                    .Where(_ => _.ClientId != clientId),
                                client with
                                {
                                    Expires = expires
                                }
                            ]
                        });
                    },
                    null,
                    cancel);
                if (result.ErrorCode == success)
                {
                    // None without a session, as SharePoint Server SE answers
                    var status = "None";
                    if (result.Session.Clients.Count > 0)
                    {
                        status = Status(result.Session);
                    }

                    data.Add(new XAttribute("CoauthStatus", status));
                }

                return result.ErrorCode;
            }
            default:
                return "InvalidArgument";
        }
    }

    async Task<string> SchemaLock(XElement? requestData, FssHttpSite site, XElement data, Cancel cancel)
    {
        var clientId = Attribute(requestData, "ClientID");
        var schemaLockId = Attribute(requestData, "SchemaLockID");
        if (clientId == null || schemaLockId == null)
        {
            return "InvalidArgument";
        }

        if (TimeoutTooLong(requestData))
        {
            return "LockRequestFail";
        }

        var timeout = CoauthSessions.Timeout(Int(requestData, "Timeout"));
        switch (Attribute(requestData, "SchemaLockRequestType"))
        {
            case "GetLock":
            case "RefreshLock":
            {
                var result = await Change(site, _ => Join(_, site, clientId, schemaLockId, timeout), null, cancel);
                if (result.ErrorCode == success)
                {
                    data.Add(new XAttribute("LockType", "SchemaLock"));
                }

                return result.ErrorCode;
            }
            case "ReleaseLock":
                return (await Change(site, _ => Leave(_, clientId, schemaLockId, false), null, cancel)).ErrorCode;
            case "ConvertToExclusive":
                return (await Change(
                    site,
                    _ => ConvertToExclusive(_, site, clientId, schemaLockId, Attribute(requestData, "ExclusiveLockID"), timeout, Bool(requestData, "ReleaseLockOnConversionToExclusiveFailure")),
                    clientId,
                    cancel)).ErrorCode;
            case "CheckLockAvailability":
                return await Availability(site, await sessions.Get(site.FileId, cancel), schemaLockId, null, cancel);
            default:
                return "InvalidArgument";
        }
    }

    async Task<string> ExclusiveLock(XElement? requestData, FssHttpSite site, XElement data, Cancel cancel)
    {
        var exclusiveLockId = Attribute(requestData, "ExclusiveLockID");
        if (exclusiveLockId == null)
        {
            return "InvalidArgument";
        }

        if (TimeoutTooLong(requestData))
        {
            return "LockRequestFail";
        }

        var timeout = Math.Clamp(Int(requestData, "Timeout") ?? 3600, 60, 120_000);
        var clientId = Attribute(requestData, "ClientID");
        var schemaLockId = Attribute(requestData, "SchemaLockID");
        switch (Attribute(requestData, "ExclusiveLockRequestType"))
        {
            case "GetLock":
            case "RefreshLock":
                return await TakeExclusiveLock(site, exclusiveLockId, timeout, cancel);
            case "ReleaseLock":
                if (!site.Access.CanWrite)
                {
                    return "FileUnauthorizedAccess";
                }

                return (await Change(
                    site,
                    _ =>
                    {
                        if (_.Clients.Count > 0 ||
                            (_.ExclusiveLockId != null && _.ExclusiveLockId != exclusiveLockId) ||
                            _.ExclusiveOfAnother(site.Access.UserId))
                        {
                            return new("FileAlreadyLockedOnServer", _);
                        }

                        // nothing to release is a Success, as SharePoint Server SE answers
                        if (_.ExclusiveLockId == null)
                        {
                            return new(success, _);
                        }

                        return new(success, _ with
                        {
                            ExclusiveLockId = null
                        });
                    },
                    null,
                    cancel)).ErrorCode;
            case "ConvertToSchemaJoinCoauth":
            case "ConvertToSchema":
            {
                if (clientId == null || schemaLockId == null)
                {
                    return "InvalidArgument";
                }

                var result = await Change(
                    site,
                    _ =>
                    {
                        if (_.ExclusiveLockId == null)
                        {
                            return new("FileNotLockedOnServer", _);
                        }

                        if (_.ExclusiveLockId != exclusiveLockId ||
                            _.Clients.Count > 0 ||
                            _.ExclusiveOfAnother(site.Access.UserId))
                        {
                            return new("FileAlreadyLockedOnServer", _);
                        }

                        return Join(_ with { ExclusiveLockId = null }, site, clientId, schemaLockId, CoauthSessions.Timeout(timeout));
                    },
                    null,
                    cancel);
                if (result.ErrorCode == success &&
                    Attribute(requestData, "ExclusiveLockRequestType") == "ConvertToSchemaJoinCoauth")
                {
                    data.Add(new XAttribute("CoauthStatus", Status(result.Session)), new XAttribute("TransitionID", site.ResourceId));
                }

                return result.ErrorCode;
            }
            case "CheckLockAvailability":
            {
                var session = await sessions.Get(site.FileId, cancel);
                if (session.Clients.Count > 0)
                {
                    return "FileAlreadyLockedOnServer";
                }

                return await Availability(site, session, null, exclusiveLockId, cancel);
            }
            default:
                return "InvalidArgument";
        }
    }

    /// <summary>
    /// Takes, or refreshes, the exclusive lock of the id, for <paramref name="timeout"/> seconds (clamped as the
    /// ExclusiveLock sub-request clamps it).
    /// </summary>
    async Task<string> TakeExclusiveLock(FssHttpSite site, string exclusiveLockId, int timeout, Cancel cancel)
    {
        timeout = Math.Clamp(timeout, 60, 120_000);
        return (await Change(
            site,
            _ =>
            {
                if (!site.Access.CanWrite)
                {
                    return new("FileUnauthorizedAccess", _);
                }

                if (_.Clients.Count > 0 ||
                    (_.ExclusiveLockId != null && _.ExclusiveLockId != exclusiveLockId) ||
                    _.ExclusiveOfAnother(site.Access.UserId))
                {
                    return new("FileAlreadyLockedOnServer", _);
                }

                return new(success, _ with
                {
                    ExclusiveLockId = exclusiveLockId,
                    ExclusiveExpires = sessions.Now.AddSeconds(timeout),
                    ExclusiveOwner = site.Access.UserName ?? site.Access.UserId,
                    ExclusiveUserId = site.Access.UserId
                });
            },
            null,
            cancel)).ErrorCode;
    }

    /// <summary>
    /// The metadata of the editor a join of the shared lock makes: the time of the join, as SharePoint Server SE adds it
    /// (docs/fsshttp-notes.md), a FILETIME in base64.
    /// </summary>
    Dictionary<string, string> JoinMetadata()
    {
        var time = new byte[8];
        BinaryPrimitives.WriteInt64LittleEndian(time, sessions.Now.ToFileTime());
        return new()
        {
            [joinCoauthTime] = Convert.ToBase64String(time)
        };
    }

    const string joinCoauthTime = "JoinCoauthTime";

    /// <summary>
    /// The editors table sub-requests (3.1.4.8): an entry per client, which a join of the shared lock also makes.
    /// Metadata is limited as in 2.3.1.23: keys of 64 bytes, values of 1,024, four pairs, besides the JoinCoauthTime of
    /// the server.
    /// </summary>
    async Task<string> EditorsTable(XElement? requestData, FssHttpSite site, Cancel cancel)
    {
        var clientId = Attribute(requestData, "ClientID");
        if (clientId == null)
        {
            return "InvalidArgument";
        }

        var timeout = CoauthSessions.Timeout(Int(requestData, "Timeout"));
        var type = Attribute(requestData, "EditorsTableRequestType");
        var key = Attribute(requestData, "Key");
        var value = Attribute(requestData, "Value") ?? "";
        return (await sessions.Change(
            site.FileId,
            _ =>
            {
                var editor = _.Editors.FirstOrDefault(_ => _.ClientId == clientId);
                var others = _.Editors.Where(_ => _.ClientId != clientId);
                switch (type)
                {
                    case "JoinEditingSession":
                    case "RefreshEditingSession":
                        if (editor == null &&
                            _.Editors.Count >= maxEditors)
                        {
                            return new("NumberOfCoauthorsReachedMax", _);
                        }

                        var asEditor = Bool(requestData, "AsEditor") && site.Access.CanWrite;
                        var entry = new SessionEditor(clientId, site.Access.UserId, site.Access.UserName, asEditor, sessions.Now.AddSeconds(timeout), editor?.Metadata ?? []);
                        return new(success, _ with
                        {
                            Editors = [.. others, entry]
                        });
                    case "LeaveEditingSession":
                        return new(success, _ with
                        {
                            Editors = others.ToList()
                        });
                    case "UpdateEditorMetadata":
                    case "RemoveEditorMetadata":
                        if (editor == null)
                        {
                            return new("EditorClientIdNotFound", _);
                        }

                        if (key == null)
                        {
                            return new("InvalidArgument", _);
                        }

                        var metadata = new Dictionary<string, string>(editor.Metadata);
                        if (type == "RemoveEditorMetadata")
                        {
                            metadata.Remove(key);
                        }
                        else
                        {
                            if (Encoding.UTF8.GetByteCount(key) > 64 ||
                                value.Length > 1368)
                            {
                                return new("EditorMetadataStringExceedsLengthLimit", _);
                            }

                            if (!metadata.ContainsKey(key) &&
                                metadata.Keys.Count(_ => _ != joinCoauthTime) >= 4)
                            {
                                return new("EditorMetadataQuotaReached", _);
                            }

                            metadata[key] = value;
                        }

                        return new(success, _ with
                        {
                            Editors =
                            [
                                .. others, editor with
                                {
                                    Metadata = metadata
                                }
                            ]
                        });
                    default:
                        return new("InvalidArgument", _);
                }
            },
            cancel)).ErrorCode;
    }

    /// <summary>
    /// The editors table as [MS-FSSHTTPB] 2.2.3.1.2.1 has the server return it: the XML of the editors, compressed with
    /// DEFLATE after the Editors Table Zip Stream Header, as a file stream ([MS-FSSHTTPD]). Its expiry times are rounded
    /// up to ten minutes, so that the polls that keep the clients in do not change it each time.
    /// </summary>
    public static byte[] EditorsTableStream(Session session)
    {
        var xml = new XElement(
            "EditorsTable",
            session.Editors.OrderBy(_ => _.ClientId, StringComparer.Ordinal).Select(_ =>
                new XElement(
                    "Editor",
                    new XElement("CacheID", _.ClientId),
                    new XElement("FriendlyName", _.UserName ?? _.UserId),
                    new XElement("LoginName", _.UserId),
                    new XElement("HasEditorPermission", XmlConvert.ToString(_.AsEditor)),
                    new XElement("Timeout", RoundUp(_.Expires).UtcTicks),
                    new XElement("Metadata", _.Metadata.OrderBy(_ => _.Key, StringComparer.Ordinal).Select(_ => new XElement(XmlConvert.EncodeLocalName(_.Key), _.Value))))));
        using var output = new MemoryStream();
        output.Write(editorsTableHeader);
        using (var deflate = new DeflateStream(output, CompressionLevel.Optimal, true))
        {
            deflate.Write(Encoding.UTF8.GetBytes(xml.ToString(SaveOptions.DisableFormatting)));
        }

        return output.ToArray();
    }

    static DateTimeOffset RoundUp(DateTimeOffset value)
    {
        var step = TimeSpan.FromMinutes(10).Ticks;
        return new((value.UtcTicks + step - 1) / step * step, TimeSpan.Zero);
    }

    /// <summary>
    /// The cell of the editors table, written again when the table changed.
    /// </summary>
    async Task<StoredCell> EditorsTableCell(string fileId, Cancel cancel)
    {
        var stream = EditorsTableStream(await sessions.Get(fileId, cancel));
        var version = Convert.ToHexString(SHA256.HashData(stream));
        var cell = await cells.Get(fileId, FssHttpCells.EditorsTablePartition, cancel);
        if (cell.State.Storage != null &&
            cell.State.FileVersion == version)
        {
            return cell;
        }

        var ids = cell.State.Ids();
        var previous = cell.State.Storage;
        if (previous != null &&
            CellSync.IsEmpty(previous))
        {
            previous = null;
        }

        var update = FileCell.Write(stream, previous, ids, ChunkingOptions.Default);
        var state = new CellState(update.Storage, ids.Guid, ids.Next, version);
        return await cells.TryUpdate(fileId, FssHttpCells.EditorsTablePartition, cell.Version, state, cancel) ??
               new StoredCell(null, state);
    }

    /// <summary>
    /// WhoAmI (3.1.4.6): the user of the token.
    /// </summary>
    static string WhoAmI(FssHttpSite site, XElement data)
    {
        data.Add(
            new XAttribute("UserName", site.Access.UserName ?? site.Access.UserId),
            new XAttribute("UserLogin", site.Access.UserId),
            new XAttribute("UserEmailAddress", ""),
            new XAttribute("UserSIPAddress", ""));
        return success;
    }

    /// <summary>
    /// ServerTime (3.1.4.7): ticks since 0001, UTC.
    /// </summary>
    string ServerTime(XElement data)
    {
        data.Add(new XAttribute("ServerTime", sessions.Now.UtcTicks));
        return success;
    }

    static XElement Property(string key, object value) =>
        new(soap + "Property", new XAttribute("Key", key), new XAttribute("Value", value));

    /// <summary>
    /// GetDocMetaInfo (3.1.4.9): some of the vti properties of the file ([MS-FPSE] 2.2.4) SharePoint returns, and of its
    /// lock. Excel opens a file read-only without vti_canmaybeedit (docs/fsshttp-notes.md).
    /// </summary>
    async Task<string> GetDocMetaInfo(FssHttpSite site, XElement data, Cancel cancel)
    {
        var file = await storage.GetFile(site.FileId, cancel);
        if (file == null)
        {
            return "FileNotExistsOrCannotBeCreated";
        }

        var modified = file.LastModified.UtcDateTime.ToString("yyyy-MM-ddTHH:mm:ss", CultureInfo.InvariantCulture);
        var canEdit = site.Access.CanWrite ? "true" : "false";
        var properties = new List<XElement>
        {
            Property("vti_title", Path.GetFileNameWithoutExtension(file.Name)),
            Property("vti_filesize", file.Size),
            Property("vti_timecreated", modified),
            Property("vti_timelastmodified", modified),
            Property("vti_author", file.OwnerId),
            Property("vti_modifiedby", site.Access.UserId),
            Property("vti_replid", $"rid:{{{site.ResourceId.ToString().ToUpperInvariant()}}}"),
            Property("vti_contenttag", $"{{{site.ResourceId.ToString().ToUpperInvariant()}}},{file.Version},1"),
            Property("vti_canmaybeedit", canEdit)
        };
        // the lock: 0 exclusive, 1 shared (the coauthoring lock)
        var session = await sessions.Get(site.FileId, cancel);
        if (session.ExclusiveLockId != null)
        {
            properties.Add(Property("vti_sourcecontrollockid", session.ExclusiveLockId.Trim('{', '}').ToLowerInvariant()));
            properties.Add(Property("vti_sourcecontrollocktype", "0"));
            if (session.ExclusiveOwner != null)
            {
                properties.Add(Property("vti_sourcecontrolcheckedoutby", session.ExclusiveOwner));
            }
        }
        else if (session.SchemaLockId != null)
        {
            properties.Add(Property("vti_sourcecontrollockid", session.SchemaLockId.Trim('{', '}').ToLowerInvariant()));
            properties.Add(Property("vti_sourcecontrollocktype", "1"));
        }

        data.Add(new XElement(soap + "DocProps", properties), new XElement(soap + "FolderProps"));
        return success;
    }

    /// <summary>
    /// LockStatus (2.3.1.50): the lock of the file, its id and its user, as Excel asks when it cannot take the exclusive
    /// lock, and names the user in its File In Use. A WOPI lock is exclusive, with neither. The LockType is a number, as
    /// SharePoint sends it (0 none, 1 shared, 2 exclusive): Excel ignores the answer with the name of the type
    /// (docs/fsshttp-notes.md).
    /// </summary>
    async Task<string> LockStatus(FssHttpSite site, XElement data, Cancel cancel)
    {
        var session = await sessions.Get(site.FileId, cancel);
        var lockId = session.ExclusiveLockId ?? session.SchemaLockId;
        if (lockId != null)
        {
            var type = "1";
            if (session.ExclusiveLockId != null)
            {
                type = "2";
            }

            data.Add(new XAttribute("LockType", type), new XAttribute("LockID", lockId.Trim('{', '}').ToLowerInvariant()));
            if (LockHolder(session) is { } holder)
            {
                data.Add(new XAttribute("LockedBy", holder));
            }

            return success;
        }

        if (await sessions.ForeignLock(site.FileId, cancel) != null)
        {
            data.Add(new XAttribute("LockType", "2"));
            return success;
        }

        data.Add(new XAttribute("LockType", "0"));
        return success;
    }

    /// <summary>
    /// FileOperation (2.3.1.34): Rename, to NewFileName, with its extension, through <see cref="IOfficeStorage.Rename"/>.
    /// As SharePoint Server SE answers (docs/fsshttp-notes.md): a file in a coauthoring session, or under a lock other
    /// than the exclusive lock of the ExclusiveLockID, is FileAlreadyLockedOnServer; a name taken is FileAlreadyExists;
    /// an illegal name is InvalidArgument; no name is a Success that renames nothing. The site of the file keeps
    /// working under its new name, the token in its path standing for the file.
    /// </summary>
    async Task<string> FileOperation(XElement? requestData, FssHttpSite site, Cancel cancel)
    {
        if (Attribute(requestData, "FileOperation") != "Rename")
        {
            return "InvalidArgument";
        }

        var newName = Attribute(requestData, "NewFileName");
        if (string.IsNullOrEmpty(newName))
        {
            return success;
        }

        if (!site.Access.CanRename)
        {
            return "FileUnauthorizedAccess";
        }

        if (!FileNames.IsLegal(newName))
        {
            return "InvalidArgument";
        }

        var session = await sessions.Get(site.FileId, cancel);
        var exclusiveLockId = Attribute(requestData, "ExclusiveLockID");
        if (session.Clients.Count > 0 ||
            (session.ExclusiveLockId != null && session.ExclusiveLockId != exclusiveLockId) ||
            await sessions.ForeignLock(site.FileId, cancel) != null)
        {
            return "FileAlreadyLockedOnServer";
        }

        try
        {
            if (await storage.Rename(site.FileId, newName, cancel) == null)
            {
                return "FileAlreadyExists";
            }

            return success;
        }
        catch (FileNotFoundException)
        {
            return "FileNotExistsOrCannotBeCreated";
        }
    }

    /// <summary>
    /// The name of the user who has the lock of the session: that of the exclusive lock, or the first editor of the
    /// shared lock; null when unknown.
    /// </summary>
    static string? LockHolder(Session session)
    {
        if (session.ExclusiveLockId != null)
        {
            return session.ExclusiveOwner;
        }

        if (session is
            {
                SchemaLockId: not null,
                Editors.Count: > 0
            })
        {
            var editor = session.Editors[0];
            return editor.UserName ?? editor.UserId;
        }

        return null;
    }

    /// <summary>
    /// AmIAlone (3.1.4.13).
    /// </summary>
    async Task<string> AmIAlone(FssHttpSite site, XElement data, Cancel cancel)
    {
        var session = await sessions.Get(site.FileId, cancel);
        // capitalized, as SharePoint Server SE answers
        var alone = session.Clients.Count <= 1 ? "True" : "False";
        data.Add(new XAttribute("AmIAlone", alone));
        return success;
    }

    /// <summary>
    /// Properties (2.3.1.48): as SharePoint Server SE answers, PropertyGet gives no values (the properties of
    /// SharePoint Online, such as RealtimeTypingEndpointUrl, are not there), and PropertyEnumerate is not supported.
    /// Office sends neither to a SharePoint Server.
    /// </summary>
    static string Properties(XElement? requestData, XElement data)
    {
        if (Attribute(requestData, "Properties") != "PropertyGet")
        {
            return "RequestNotSupported";
        }

        data.Add(new XElement(soap + "PropertyValues"));
        return success;
    }

    /// <summary>
    /// A lock timeout of 120000 seconds or more: the spec's range ends at 120000 ([MS-FSSHTTP] 2.3.3.2), but SharePoint
    /// Server SE fails a join or a lock from 120000 on with LockRequestFail.
    /// </summary>
    static bool TimeoutTooLong(XElement? requestData) => Int(requestData, "Timeout") >= 120_000;
}
