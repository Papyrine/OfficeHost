using System.IO.Compression;

/// <summary>
/// The file and user a cell storage request is for, from the token in its path, and the URL of its site.
/// </summary>
record FssHttpSite(OfficeAccess Access, string WebUrl)
{
    public string FileId => Access.FileId;

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
class CellStorageService(IOfficeStorage storage, FssHttpCells cells, CoauthSessions sessions)
{
    static XNamespace soap = SoapMessage.Soap;
    // E_FAIL, the HResult of SharePoint with an error code (MS-FSSHTTP 4.1.2)
    const long failHResult = 2147500037;
    const string success = "Success";

    // the Editors Table Zip Stream Header ([MS-FSSHTTPB] 2.2.3.1.2.1.1)
    static byte[] editorsTableHeader = [0x1A, 0x5A, 0x3A, 0x30, 0x00, 0x00, 0x00, 0x00];

    public async Task<XElement> Execute(SoapMessage message, FssHttpSite site, MtomResponse mtom, Cancel cancel)
    {
        var body = message.Document.Root?.Element(SoapMessage.Envelope + "Body") ??
                   throw new InvalidDataException("No SOAP body");
        var responses = new List<XElement>();
        foreach (var request in body.Descendants(soap + "Request"))
        {
            responses.Add(await Request(message, request, site, mtom, cancel));
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
            // SharePoint names the user who has the lock in the ErrorMessage, which Excel shows in its File In Use
            if (errorCode == "FileAlreadyLockedOnServer" &&
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
    /// is an MTOM part; a Put Changes that fails (a coherency failure) fails in it, with the error code Success.
    /// </summary>
    async Task<string> Cell(SoapMessage message, XElement? requestData, FssHttpSite site, XElement data, MtomResponse mtom, Cancel cancel)
    {
        if (requestData == null ||
            message.Binary(requestData) is not { } binary)
        {
            return success;
        }

        var file = await storage.GetFile(site.FileId, cancel);
        if (file == null)
        {
            return "CellRequestFail";
        }

        var request = FssHttpBRequest.Read(binary);
        var partitionAttribute = (string?) requestData.Attribute("PartitionID");
        var defaultPartition = Guid.Empty;
        if (partitionAttribute != null)
        {
            defaultPartition = Guid.Parse(partitionAttribute);
        }

        var subResponses = new List<SubResponse>();
        var elements = new Dictionary<ExtendedGuid, DataElement>();
        var package = request.DataElements?.Elements ?? [];
        foreach (var subRequest in request.SubRequests)
        {
            var partition = subRequest.TargetPartition ?? defaultPartition;
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
                    subResponses.Add(new AllocateExtendedGuidRangeResponse(allocate.RequestId, Guid.NewGuid(), 1, allocate.Count + 1));
                    break;
            }
        }

        var fssHttpB = new FssHttpBResponse(subResponses, new(elements.Values.ToList()), ProtocolVersion: 13).ToArray();
        if (requestData.Attribute("GetFileProps")?.Value is "true" or "True")
        {
            file = await storage.GetFile(site.FileId, cancel) ?? file;
            data.Add(
                new XAttribute("Etag", Etag(site, file)),
                new XAttribute("CreateTime", file.LastModified.ToFileTime()),
                new XAttribute("LastModifiedTime", file.LastModified.ToFileTime()),
                new XAttribute("ModifiedBy", site.Access.UserName ?? site.Access.UserId),
                new XAttribute("HaveOnlyDemotionChanges", "False"));
        }

        data.Add(mtom.Include(fssHttpB));
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
            !await sessions.CanWrite(site.FileId, lockId, cancel))
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

            var pending = await cells.TryUpdate(site.FileId, partition, cell.Version, outcome.State with { FileVersion = null }, cancel);
            if (pending == null)
            {
                continue;
            }

            await cells.WriteFile(site.FileId, pending, bytes, cancel);
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

        return int.Parse(value, CultureInfo.InvariantCulture);
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
        var clients = session.Clients.Where(_ => _.ClientId != clientId).Append(new(clientId, expires, timeout)).ToList();
        var editor = session.Editors.FirstOrDefault(_ => _.ClientId == clientId);
        var editors = session.Editors.Where(_ => _.ClientId != clientId)
            .Append(new(clientId, site.Access.UserId, site.Access.UserName, true, expires, editor?.Metadata ?? []))
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

    SessionResult ConvertToExclusive(Session session, string clientId, string schemaLockId, string? exclusiveLockId, int timeout, bool release)
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
            Editors = session.Editors
        });
    }

    static string Availability(Session session, string? schemaLockId, string? exclusiveLockId)
    {
        if ((session.ExclusiveLockId != null && session.ExclusiveLockId != exclusiveLockId) ||
            (session.SchemaLockId != null && session.SchemaLockId != schemaLockId))
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
                        new XAttribute("CoauthStatus", Status(result.Session)),
                        new XAttribute("TransitionID", site.ResourceId));
                }

                return result.ErrorCode;
            }
            case "ExitCoauthoring":
                return (await Change(site, _ => Leave(_, clientId, schemaLockId, true), null, cancel)).ErrorCode;
            case "ConvertToExclusive":
                return (await Change(
                    site,
                    _ => ConvertToExclusive(_, clientId, schemaLockId, Attribute(requestData, "ExclusiveLockID"), timeout, Bool(requestData, "ReleaseLockOnConversionToExclusiveFailure")),
                    clientId,
                    cancel)).ErrorCode;
            case "CheckLockAvailability":
                return Availability(await sessions.Get(site.FileId, cancel), schemaLockId, null);
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
                    _ => ConvertToExclusive(_, clientId, schemaLockId, Attribute(requestData, "ExclusiveLockID"), timeout, Bool(requestData, "ReleaseLockOnConversionToExclusiveFailure")),
                    clientId,
                    cancel)).ErrorCode;
            case "CheckLockAvailability":
                return Availability(await sessions.Get(site.FileId, cancel), schemaLockId, null);
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

        var timeout = Math.Clamp(Int(requestData, "Timeout") ?? 3600, 60, 120_000);
        var clientId = Attribute(requestData, "ClientID");
        var schemaLockId = Attribute(requestData, "SchemaLockID");
        switch (Attribute(requestData, "ExclusiveLockRequestType"))
        {
            case "GetLock":
            case "RefreshLock":
                return (await Change(
                    site,
                    _ =>
                    {
                        if (!site.Access.CanWrite)
                        {
                            return new("FileUnauthorizedAccess", _);
                        }

                        if (_.Clients.Count > 0 ||
                            (_.ExclusiveLockId != null && _.ExclusiveLockId != exclusiveLockId))
                        {
                            return new("FileAlreadyLockedOnServer", _);
                        }

                        return new(success, _ with
                        {
                            ExclusiveLockId = exclusiveLockId,
                            ExclusiveExpires = sessions.Now.AddSeconds(timeout),
                            ExclusiveOwner = site.Access.UserName ?? site.Access.UserId
                        });
                    },
                    null,
                    cancel)).ErrorCode;
            case "ReleaseLock":
                return (await Change(
                    site,
                    _ =>
                    {
                        if (_.Clients.Count > 0 ||
                            (_.ExclusiveLockId != null && _.ExclusiveLockId != exclusiveLockId))
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
                            _.Clients.Count > 0)
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

                return Availability(session, null, exclusiveLockId);
            }
            default:
                return "InvalidArgument";
        }
    }

    /// <summary>
    /// The editors table sub-requests (3.1.4.8): an entry per client, which a join of the shared lock also makes.
    /// Metadata is limited as in 2.3.1.23: keys of 64 bytes, values of 1,024, four pairs.
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
                                metadata.Count >= 4)
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
        data.Add(new XAttribute("AmIAlone", session.Clients.Count <= 1 ? "true" : "false"));
        return success;
    }

}
