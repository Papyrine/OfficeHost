/// <summary>
/// The FSSHTTPB codec on published and captured bytes (<see cref="FssHttpBVectors"/>): each is read, checked against what
/// the spec says it holds, and written back byte for byte.
/// </summary>
public class FssHttpBVectorTests
{
    public static IEnumerable<string> Names() => FssHttpBVectors.All().Select(_ => _.Name);

    [Test]
    [MethodDataSource(nameof(Names))]
    public async Task RoundTrip(string name)
    {
        var (_, bytes, parse) = FssHttpBVectors.All().Single(_ => _.Name == name);
        var written = FssHttpB.Write(parse(bytes));
        await Assert.That(FssHttpB.Hex(written)).IsEqualTo(FssHttpB.Hex(bytes));
    }

    /// <summary>
    /// [MS-FSSHTTPB] 4.1: a Query Changes request with a user agent GUID and version, no scoping cell, a maximum of
    /// 3670016 bytes of data elements, empty knowledge, and an empty Data Element Package.
    /// </summary>
    [Test]
    public async Task Spec_QueryChangesRequest()
    {
        var request = FssHttpBRequest.Read(FssHttpB.Hex(FssHttpBVectors.SpecQueryChangesRequest));
        await Assert.That(request.ProtocolVersion).IsEqualTo((ushort)12);
        await Assert.That(request.MinimumVersion).IsEqualTo((ushort)11);
        await Assert.That(request.UserAgent!.Guid).IsEqualTo(FssHttpB.Guid("E731B87E-DD45-44AA-AB80-0C75FBD1530E"));
        await Assert.That(request.UserAgent.Client).IsNull();
        await Assert.That(request.UserAgent.Version).IsEqualTo(0x0FA127C4u);
        var query = (QueryChangesRequest)request.SubRequests.Single();
        await Assert.That(query.RequestId).IsEqualTo(1ul);
        await Assert.That(query.Priority).IsEqualTo(0ul);
        await Assert.That(query.Flags).IsEqualTo((ushort)0);
        await Assert.That(query.AllowFragments).IsFalse();
        await Assert.That(query.Arguments!.IncludeStorageManifest).IsTrue();
        await Assert.That(query.Arguments.IncludeCellChanges).IsTrue();
        await Assert.That(query.Arguments.Cell).IsEqualTo(CellId.Null);
        await Assert.That(query.MaxDataElements).IsEqualTo(3670016ul);
        await Assert.That(query.Knowledge!.Items.Count).IsEqualTo(0);
        await Assert.That(request.DataElements!.Elements.Count).IsEqualTo(0);
    }

    /// <summary>
    /// [MS-FSSHTTPB] 4.2: a Query Changes sub-response with the Storage Index, and knowledge of two cell ranges and a
    /// waterline.
    /// </summary>
    [Test]
    public async Task Spec_QueryChangesSubResponse()
    {
        var bytes = FssHttpB.Hex(FssHttpBVectors.SpecQueryChangesSubResponse);
        var response = (QueryChangesResponse)FssHttpB.Read(bytes, (ref FssHttpBReader _) => SubResponse.Read(ref _));
        await Assert.That(response.RequestId).IsEqualTo(1ul);
        await Assert.That(response.StorageIndex).IsEqualTo(new(FssHttpB.Guid("A00D98FD-40FD-4D99-930A-6322D7689136"), 1));
        await Assert.That(response.Partial).IsFalse();
        var cell = (CellKnowledge)response.Knowledge.Items[0];
        var first = (CellKnowledgeRange)cell.Items[0];
        await Assert.That(first.Guid).IsEqualTo(FssHttpB.Guid("E20A9380-FD55-BCA5-9037-451C9D86E949"));
        await Assert.That(first.From).IsEqualTo(0ul);
        await Assert.That(first.To).IsEqualTo(73507ul);
        var second = (CellKnowledgeRange)cell.Items[1];
        await Assert.That(second.Guid).IsEqualTo(FssHttpB.Guid("1DF56C7F-02AA-435A-9037-451C9D86E949"));
        // the spec text says 75503: the bytes FC F8 08 are 73503
        await Assert.That(second.To).IsEqualTo(73503ul);
        var waterline = ((WaterlineKnowledge)response.Knowledge.Items[1]).Entries.Single();
        await Assert.That(waterline.CellStorage).IsEqualTo(new(FssHttpB.Guid("1DF56C7F-02AA-435A-9037-451C9D86E949"), 1));
        await Assert.That(waterline.Waterline).IsEqualTo(73503ul);
    }

    /// <summary>
    /// [MS-FSSHTTPB] 4.4: a Put Changes response with cell and content tag knowledge, and no Put Changes response
    /// object.
    /// </summary>
    [Test]
    public async Task Spec_PutChangesResponse()
    {
        var response = FssHttpBResponse.Read(FssHttpB.Hex(FssHttpBVectors.SpecPutChangesResponse));
        await Assert.That(response.Error).IsNull();
        await Assert.That(response.DataElements).IsNull();
        var put = (PutChangesResponse)response.SubResponses.Single();
        await Assert.That(put.HasResponseObject).IsFalse();
        var ranges = ((CellKnowledge)put.Knowledge.Items[0]).Items.Cast<CellKnowledgeRange>().ToList();
        await Assert.That(ranges[0].Guid).IsEqualTo(FssHttpB.Guid("92699222-AD46-B353-9489-C24F5ACFA09A"));
        await Assert.That(ranges[0].To).IsEqualTo(116ul);
        await Assert.That(ranges[1].Guid).IsEqualTo(FssHttpB.Guid("6D966DDD-52B9-4CAC-9489-C24F5ACFA09A"));
        await Assert.That(ranges[1].To).IsEqualTo(111ul);
        var tag = ((ContentTagKnowledge)put.Knowledge.Items[1]).Entries.Single();
        await Assert.That(tag.Blob).IsEqualTo(new(FssHttpB.Guid("37410BF9-D16F-4499-A6C3-27232EDCA711"), 1));
        await Assert.That(FssHttpB.Hex(tag.ClockData.Span)).IsEqualTo("33000000");
    }

    /// <summary>
    /// [MS-FSSHTTPB] 4.3: a Put Changes request, of the Storage Index 052E2E8E-C0D1-4886-9C51-29D661714F67,1, with no
    /// expected Storage Index and the flags 0x48, and its Storage Manifest, Cell Manifest, Revision Manifest and Storage
    /// Index.
    /// </summary>
    [Test]
    public async Task Spec_PutChangesRequest()
    {
        var request = FssHttpBRequest.Read(FssHttpBVectors.SpecPutChangesRequest);
        var put = (PutChangesRequest)request.SubRequests.Single();
        var storageIndexId = new ExtendedGuid(FssHttpB.Guid("052E2E8E-C0D1-4886-9C51-29D661714F67"), 1);
        await Assert.That(put.StorageIndex).IsEqualTo(storageIndexId);
        await Assert.That(put.ExpectedStorageIndex).IsEqualTo(ExtendedGuid.Null);
        await Assert.That(put.Flags).IsEqualTo((byte)0x48);
        await Assert.That(put.FavorCoherencyFailureOverNotFound).IsTrue();
        await Assert.That(put.ImplyNullExpectedIfNoMapping).IsFalse();
        var elements = request.DataElements!.Elements;
        await Assert.That(elements.Select(_ => _.Type).ToList()).IsEquivalentTo(
            [DataElementType.StorageManifest, DataElementType.CellManifest, DataElementType.RevisionManifest, DataElementType.StorageIndex]);
        var storageIndex = (StorageIndex)elements[3];
        await Assert.That(storageIndex.Id).IsEqualTo(storageIndexId);
        await Assert.That(storageIndex.Entries.Count).IsEqualTo(3);
    }

    /// <summary>
    /// [MS-FSSHTTPB] 4.3.3: a Storage Manifest with its schema GUID and one root.
    /// </summary>
    [Test]
    public async Task Spec_StorageManifest()
    {
        var manifest = (StorageManifest)FssHttpB.Read(FssHttpB.Hex(FssHttpBVectors.SpecStorageManifest), (ref FssHttpBReader _) => DataElement.Read(ref _));
        await Assert.That(manifest.Id).IsEqualTo(new(FssHttpB.Guid("D730FA99-122C-4288-B722-0A125CFDA7E5"), 1));
        await Assert.That(manifest.SerialNumber).IsEqualTo(new(FssHttpB.Guid("5430AF47-6E71-409B-9806-707E818DC102"), 0x32));
        await Assert.That(manifest.SchemaId).IsEqualTo(FssHttpB.Guid("0EB93394-571D-41E9-AAD3-880D92D31955"));
        var root = manifest.Roots.Single();
        await Assert.That(root.RootId).IsEqualTo(new(FssHttpB.Guid("84DEFAB9-AAA3-4A0D-A3A8-520C77AC7073"), 2));
        await Assert.That(root.Cell).IsEqualTo(
            new(
                new(FssHttpB.Guid("84DEFAB9-AAA3-4A0D-A3A8-520C77AC7073"), 1),
                new(FssHttpB.Guid("6F2A4665-42C8-46C7-BAB4-E28FDCE1E32B"), 1)));
    }

    /// <summary>
    /// [MS-FSSHTTPB] 4.3.4: a Cell Manifest, with the 10-bit form of the Extended GUID (the value 0x31).
    /// </summary>
    [Test]
    public async Task Spec_CellManifest()
    {
        var manifest = (CellManifest)FssHttpB.Read(FssHttpB.Hex(FssHttpBVectors.SpecCellManifest), (ref FssHttpBReader _) => DataElement.Read(ref _));
        await Assert.That(manifest.Id).IsEqualTo(new(FssHttpB.Guid("2C0BFC8E-9B04-4C61-AB49-4845E603ECA0"), 0x31));
        await Assert.That(manifest.CurrentRevision).IsEqualTo(new(FssHttpB.Guid("7128FE3A-DCBE-4301-BD84-716C456C808A"), 1));
    }

    /// <summary>
    /// [MS-FSSHTTPB] 4.3.5: a Revision Manifest with no base revision, a root whose object has the 32-bit form of the
    /// Extended GUID, and an Object Group. The spec text swaps the fields: the bytes are the revision, the null base
    /// revision, then the root declare.
    /// </summary>
    [Test]
    public async Task Spec_RevisionManifest()
    {
        var manifest = (RevisionManifest)FssHttpB.Read(FssHttpB.Hex(FssHttpBVectors.SpecRevisionManifest), (ref FssHttpBReader _) => DataElement.Read(ref _));
        await Assert.That(manifest.RevisionId).IsEqualTo(new(FssHttpB.Guid("7128FE3A-DCBE-4301-BD84-716C456C808A"), 1));
        await Assert.That(manifest.BaseRevisionId).IsEqualTo(ExtendedGuid.Null);
        var root = manifest.Roots.Single();
        await Assert.That(root.RootId).IsEqualTo(new(FssHttpB.Guid("84DEFAB9-AAA3-4A0D-A3A8-520C77AC7073"), 2));
        await Assert.That(root.ObjectId.Value).IsEqualTo(0xDE0C3813u);
        await Assert.That(manifest.ObjectGroups.Single()).IsEqualTo(new(FssHttpB.Guid("2C0BFC8E-9B04-4C61-AB49-4845E603ECA0"), 1));
    }

    /// <summary>
    /// [MS-FSSHTTPB] 4.3.6: a Storage Index with a manifest, a cell and a revision mapping.
    /// </summary>
    [Test]
    public async Task Spec_StorageIndex()
    {
        var index = (StorageIndex)FssHttpB.Read(FssHttpB.Hex(FssHttpBVectors.SpecStorageIndex), (ref FssHttpBReader _) => DataElement.Read(ref _));
        var manifest = (StorageIndexManifestMapping)index.Entries[0];
        await Assert.That(manifest.MappingId).IsEqualTo(new(FssHttpB.Guid("D730FA99-122C-4288-B722-0A125CFDA7E5"), 1));
        await Assert.That(manifest.MappingSerialNumber.Value).IsEqualTo(0x3Eul);
        var cell = (StorageIndexCellMapping)index.Entries[1];
        await Assert.That(cell.MappingId).IsEqualTo(new(FssHttpB.Guid("2C0BFC8E-9B04-4C61-AB49-4845E603ECA0"), 0x31));
        var revision = (StorageIndexRevisionMapping)index.Entries[2];
        await Assert.That(revision.Revision).IsEqualTo(new(FssHttpB.Guid("7128FE3A-DCBE-4301-BD84-716C456C808A"), 1));
        await Assert.That(revision.MappingId).IsEqualTo(new(FssHttpB.Guid("DFD1A905-9B9C-422E-B259-817AF3511454"), 1));
    }

    /// <summary>
    /// SharePoint Server SE's Query Access answer: read and write allowed, and the access response of type 0xA4 that the
    /// spec does not define, E_ACCESSDENIED chained to E_ACCESSDENIED.
    /// </summary>
    [Test]
    public async Task SharePoint_QueryAccessResponse()
    {
        var response = FssHttpBResponse.Read(Convert.FromBase64String(FssHttpBVectors.SharePointQueryAccessResponse));
        var access = (QueryAccessResponse)response.SubResponses.Single();
        await Assert.That(access.ReadAccess).IsEqualTo(ResponseError.Success);
        await Assert.That(access.WriteAccess).IsEqualTo(ResponseError.Success);
        var denied = new ResponseError(ResponseErrorType.HResult, 0x80070005);
        await Assert.That(access.ExtraAccess).IsEqualTo(denied with { Chained = denied });
    }

    /// <summary>
    /// [MS-FSSHTTP] 4.1.2: the responses of the three Cell sub-requests of a file open, each with a Data Element Package
    /// and a Query Changes sub-response.
    /// </summary>
    [Test]
    public async Task FssHttp_Responses()
    {
        foreach (var base64 in FssHttpBVectors.FssHttpResponses)
        {
            var response = FssHttpBResponse.Read(Convert.FromBase64String(base64));
            await Assert.That(response.DataElements!.Elements.Count).IsGreaterThan(0);
            var query = (QueryChangesResponse)response.SubResponses.Single();
            await Assert.That(query.StorageIndex.IsNull).IsFalse();
        }
    }

    /// <summary>
    /// Word 16.0.20326: a user agent with both the GUID and the client and platform, protocol version 14, and Query
    /// Changes flags of two bytes, whose second (reserved) byte varies from request to request.
    /// </summary>
    [Test]
    public async Task Word_QueryChanges()
    {
        var request = FssHttpBRequest.Read(Convert.FromBase64String(FssHttpBVectors.WordRequests[0]));
        await Assert.That(request.ProtocolVersion).IsEqualTo((ushort)14);
        await Assert.That(request.UserAgent!.Guid).IsEqualTo(FssHttpB.Guid("1984108C-4B93-4EEB-B320-919432D6E593"));
        await Assert.That(request.UserAgent.Client).IsEqualTo("msword");
        await Assert.That(request.UserAgent.Platform).IsEqualTo("win");
        var query = (QueryChangesRequest)request.SubRequests.Single();
        await Assert.That(query.FlagsLength).IsEqualTo(2);
        await Assert.That(query.Flags).IsEqualTo((ushort)0xAC00);
        await Assert.That(query.UserContentEquivalentVersionOk).IsFalse();
        await Assert.That(query.MaxDataElements).IsEqualTo(3670016ul);
    }

    /// <summary>
    /// Word 16.0.20326: the request of the file properties has the cell roundtrip options.
    /// </summary>
    [Test]
    public async Task Word_FileProperties()
    {
        var request = FssHttpBRequest.Read(Convert.FromBase64String(FssHttpBVectors.WordRequests[2]));
        await Assert.That(request.CellRoundtripOptions).IsEqualTo((byte)1);
    }

    /// <summary>
    /// Word 16.0.20326: a lone Query Access sub-request comes with no user agent, though 2.2.2 shows it as required.
    /// </summary>
    [Test]
    public async Task Word_QueryAccess()
    {
        var request = FssHttpBRequest.Read(Convert.FromBase64String(FssHttpBVectors.WordRequests[3]));
        await Assert.That(request.UserAgent).IsNull();
        await Assert.That(request.DataElements).IsNull();
        var access = (QueryAccessRequest)request.SubRequests.Single();
        await Assert.That(access.RequestId).IsEqualTo(1ul);
    }
}
