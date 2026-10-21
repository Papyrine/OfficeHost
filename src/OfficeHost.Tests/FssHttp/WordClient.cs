using System.Xml.Linq;
using Microsoft.AspNetCore.WebUtilities;
using Microsoft.Net.Http.Headers;

/// <summary>
/// A sub-response of the cell storage service, with the binary data of a Cell sub-response.
/// </summary>
record SubResult(string Token, string ErrorCode, XElement Data, byte[]? Binary, string? ErrorMessage = null)
{
    public FssHttpBResponse FssHttpB => FssHttpBResponse.Read(Binary!);

    public string? Attribute(string name) => (string?) Data.Attribute(name);
}

/// <summary>
/// A client of the desktop site of a file that does what Word does (docs/fsshttp-spike.md, docs/fsshttp-notes.md): SOAP
/// requests, MTOM answers, and a cell storage of its own for the content, which it queries with its knowledge and puts
/// with the Storage Index it expects under an id of its own, as Word does.
/// </summary>
class WordClient(TestServerHost host, string accessToken, string? clientId = null)
{
    public const string SchemaLockId = "{29358EC1-E813-4793-8E70-ED0344E7B73C}";

    static XNamespace soap = "http://schemas.microsoft.com/sharepoint/soap/";
    static XNamespace envelope = "http://schemas.xmlsoap.org/soap/envelope/";

    IdAllocator ids = new(Guid.NewGuid());

    public string ClientId { get; } = clientId ?? $"{{{Guid.NewGuid().ToString().ToUpperInvariant()}}}";

    public string Site => $"/fsshttp/{accessToken}";

    public CellStorage? Content { get; private set; }

    public Knowledge Knowledge { get; private set; } = Knowledge.Empty;

    public byte[] File => FileCell.Read(Content!).Root.ToArray();

    public static XElement SubRequest(string type, string token, XElement? data = null, string? dependsOn = null, string? dependencyType = null)
    {
        var element = new XElement(soap + "SubRequest", new XAttribute("Type", type), new XAttribute("SubRequestToken", token));
        if (dependsOn != null)
        {
            element.Add(new XAttribute("DependsOn", dependsOn), new XAttribute("DependencyType", dependencyType ?? "OnExecute"));
        }

        if (data != null)
        {
            element.Add(data);
        }

        return element;
    }

    public static XElement Data(params object[] content) => new(soap + "SubRequestData", content);

    public static XElement Cell(string token, byte[] fssHttpB, params XAttribute[] attributes) =>
        SubRequest("Cell", token, Data(attributes, new XAttribute("BinaryDataSize", fssHttpB.Length), Convert.ToBase64String(fssHttpB)));

    public static XElement Envelope(params XElement[] subRequests) =>
        new(
            envelope + "Envelope",
            new XElement(
                envelope + "Body",
                new XElement(soap + "RequestVersion", new XAttribute("Version", "2"), new XAttribute("MinorVersion", "2")),
                new XElement(
                    soap + "RequestCollection",
                    new XAttribute("CorrelationId", Guid.NewGuid()),
                    new XElement(
                        soap + "Request",
                        new XAttribute("Url", "https://localhost/fsshttp/doc.docx"),
                        new XAttribute("RequestToken", "1"),
                        subRequests))));

    public async Task<HttpResponseMessage> Post(string path, HttpContent content)
    {
        using var request = new HttpRequestMessage(HttpMethod.Post, $"{Site}/{path}");
        request.Content = content;
        request.Headers.Add("SOAPAction", "\"http://schemas.microsoft.com/sharepoint/soap/ICellStorages/ExecuteCellStorageRequest\"");
        return await host.Client.SendAsync(request);
    }

    /// <summary>
    /// Sends the sub-requests in one Request, and reads the SubResponses of the MTOM answer.
    /// </summary>
    public async Task<List<SubResult>> Send(params XElement[] subRequests)
    {
        var content = new StringContent(Envelope(subRequests).ToString(), Encoding.UTF8, "text/xml");
        using var response = await Post("_vti_bin/cellstorage.svc/CellStorageService", content);
        if (!response.IsSuccessStatusCode)
        {
            throw new($"The cell storage service answered {(int) response.StatusCode}");
        }

        return await Read(response);
    }

    public static async Task<List<SubResult>> Read(HttpResponseMessage response)
    {
        var type = MediaTypeHeaderValue.Parse(response.Content.Headers.ContentType!.ToString());
        var reader = new MultipartReader(HeaderUtilities.RemoveQuotes(type.Boundary).Value!, await response.Content.ReadAsStreamAsync());
        XDocument? document = null;
        var parts = new Dictionary<string, byte[]>();
        while (await reader.ReadNextSectionAsync() is { } section)
        {
            using var buffer = new MemoryStream();
            await section.Body.CopyToAsync(buffer);
            if (document == null)
            {
                document = XDocument.Parse(Encoding.UTF8.GetString(buffer.ToArray()));
                continue;
            }

            parts[section.Headers!["Content-ID"].ToString().Trim('<', '>')] = buffer.ToArray();
        }

        XNamespace xop = "http://www.w3.org/2004/08/xop/include";
        return document!.Descendants(soap + "SubResponse")
            .Select(_ =>
            {
                var data = _.Element(soap + "SubResponseData")!;
                var href = (string?) data.Element(xop + "Include")?.Attribute("href");
                byte[]? binary = null;
                if (href != null)
                {
                    binary = parts[href["cid:".Length..]];
                }

                return new SubResult((string) _.Attribute("SubRequestToken")!, (string) _.Attribute("ErrorCode")!, data, binary, (string?) _.Attribute("ErrorMessage"));
            })
            .ToList();
    }

    public static byte[] Request(params SubRequest[] subRequests) =>
        new FssHttpBRequest(new(Guid.NewGuid(), "msword", "win", 1), subRequests, new([])).ToArray();

    public static byte[] Query(Knowledge? knowledge = null, ulong max = 3670016) =>
        Request(new QueryChangesRequest(1, 0, 2, new(3, CellId.Null), max, Knowledge: knowledge ?? Knowledge.Empty));

    public XElement Join(string token = "1", int timeout = 2580) =>
        SubRequest("Coauth", token, Data(CoauthAttributes("JoinCoauthoring"), new XAttribute("Timeout", timeout), new XAttribute("AllowFallbackToExclusive", "true")));

    public XAttribute[] CoauthAttributes(string type) =>
    [
        new("CoauthRequestType", type),
        new("SchemaLockID", SchemaLockId),
        new("ClientID", ClientId)
    ];

    public async Task<SubResult> Coauth(string type, params XAttribute[] attributes) =>
        (await Send(SubRequest("Coauth", "1", Data(CoauthAttributes(type), attributes)))).Single();

    /// <summary>
    /// Queries the content with the knowledge of the client, and adds what comes to its cell storage.
    /// </summary>
    public async Task<QueryChangesResponse> Sync()
    {
        var result = (await Send(Cell("1", Query(Knowledge), new XAttribute("GetFileProps", "true")))).Single();
        if (result.ErrorCode != "Success")
        {
            throw new($"Query Changes failed: {result.ErrorCode}");
        }

        var response = result.FssHttpB;
        var query = (QueryChangesResponse) response.SubResponses.Single();
        var elements = response.DataElements?.Elements ?? [];
        Content = Content == null
            ? CellStorage.Create(elements, query.StorageIndex)
            : Content.With(elements, query.StorageIndex);
        Knowledge = query.Knowledge;
        return query;
    }

    /// <summary>
    /// Puts a new version of the file as Word does: the new data elements with null Serial Numbers, a Storage Index of a
    /// new id, and the Storage Index it expects, under a new id too, with Serial Numbers of its own for the mappings.
    /// With <paramref name="exclusive"/>, the lock id is that of an exclusive lock, as Excel saves: BypassLockID alone.
    /// </summary>
    public async Task<SubResponse> Save(byte[] file, string? lockId = SchemaLockId, bool exclusive = false)
    {
        var update = FileCell.Write(file, Content, ids, ChunkingOptions.Default);
        var index = update.Storage.StorageIndex with
        {
            Id = ids.NewId()
        };
        var expected = Content!.StorageIndex with
        {
            Id = ids.NewId(),
            Entries = Content.StorageIndex.Entries.Select(Remap).ToList()
        };
        var package = update.Added
            .Where(_ => _ is not StorageIndex)
            .Append(index)
            .Append(expected)
            .Select(_ => _ with { SerialNumber = SerialNumber.Null })
            .ToList();
        var put = new PutChangesRequest(1, index.Id, expected.Id, 0x49, new byte[3], new byte[] { 4, 0, 0 }, ClientKnowledge: Knowledge);
        var request = new FssHttpBRequest(new(Guid.NewGuid(), "msword", "win", 1), [put], new(package)).ToArray();
        var attributes = new List<XAttribute>
        {
            new("Coalesce", "true"),
            new("CoauthVersioning", "true"),
            new("GetFileProps", "true")
        };
        if (lockId != null)
        {
            attributes.Add(new("BypassLockID", lockId));
            if (!exclusive)
            {
                attributes.Add(new("SchemaLockID", lockId));
            }
        }

        var result = (await Send(Cell("1", request, attributes.ToArray()))).Single();
        if (result.ErrorCode != "Success")
        {
            return new FailedSubResponse(1, RequestType.PutChanges, new(ResponseErrorType.Protocol, 0, result.ErrorCode));
        }

        var response = result.FssHttpB.SubResponses.Single();
        if (response is PutChangesResponse succeeded)
        {
            Content = update.Storage.With([index], index.Id);
            Knowledge = succeeded.Knowledge;
        }

        return response;
    }

    StorageIndexEntry Remap(StorageIndexEntry entry) =>
        entry switch
        {
            StorageIndexManifestMapping value => value with { MappingSerialNumber = ids.NewSerialNumber() },
            StorageIndexCellMapping value => value with { MappingSerialNumber = ids.NewSerialNumber() },
            StorageIndexRevisionMapping value => value with { MappingSerialNumber = ids.NewSerialNumber() },
            _ => entry
        };
}
