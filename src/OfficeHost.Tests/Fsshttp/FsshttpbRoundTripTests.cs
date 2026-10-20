using System.Runtime.InteropServices;

/// <summary>
/// Every structure of the FSSHTTPB codec, with and without its optional parts: written, read back to the same model, and
/// written again to the same bytes.
/// </summary>
public class FsshttpbRoundTripTests
{
    static readonly Guid guid = Fsshttpb.Guid("84DEFAB9-AAA3-4A0D-A3A8-520C77AC7073");
    static readonly Guid other = Fsshttpb.Guid("6F2A4665-42C8-46C7-BAB4-E28FDCE1E32B");

    static ExtendedGuid Id(uint value) => new(guid, value);

    static SerialNumber Serial(ulong value) => new(other, value);

    static CellId Cell(uint value) => new(Id(value), new(other, value));

    static byte[] Data(int length) => Enumerable.Range(0, length).Select(_ => (byte) _).ToArray();

    static List<DataElement> Elements() =>
    [
        new StorageIndex(Id(1), Serial(1), [
            new StorageIndexManifestMapping(Id(2), Serial(2)),
            new StorageIndexCellMapping(Cell(3), Id(4), Serial(3)),
            new StorageIndexRevisionMapping(Id(5), Id(6), SerialNumber.Null)
        ]),
        new StorageIndex(Id(1), Serial(1), []),
        new StorageManifest(Id(0x400), Serial(5), other, [new(Id(1), Cell(1)), new(Id(2), CellId.Null)]),
        new StorageManifest(Id(0x400), Serial(5), other, []),
        new CellManifest(Id(0x20000), SerialNumber.Null, Id(9)),
        new RevisionManifest(Id(10), Serial(10), Id(11), Id(12), [new(Id(1), Id(0xDE0C3813))], [Id(13), Id(14)]),
        new RevisionManifest(Id(10), Serial(10), Id(11), ExtendedGuid.Null, [], []),
        new ObjectGroup(
            Id(20),
            Serial(20),
            new DataElementHash(1, Data(20)),
            [new ObjectDeclaration(Id(21), 1, 600, 2, 1), new ObjectDataBlobDeclaration(Id(22), Id(23), 1, 0, 0)],
            [1, 0x4000],
            [
                new ObjectData([Id(24), ExtendedGuid.Null], [Cell(1)], Data(600)),
                new ObjectExcludedData([], [], 1000),
                new ObjectDataBlobReference([Id(25)], [], Id(23))
            ]),
        new ObjectGroup(Id(20), Serial(20), null, [], null, []),
        new ObjectGroup(Id(20), Serial(20), null, [], [], []),
        new DataElementFragment(Id(30), Serial(30), Id(31), 100000, new(32768, 40000), Data(40000)),
        new ObjectDataBlob(Id(32), Serial(32), Data(127)),
        new ObjectDataBlob(Id(33), Serial(33), Data(70000))
    ];

    static Knowledge FullKnowledge()
    {
        var unknown = Fsshttpb.Bytes((ref FsshttpbWriter _) =>
        {
            _.WriteStart(StreamObjectType.ObjectDataBlob, 2);
            _.WriteBytes([1, 2]);
            _.WriteStart(StreamObjectType.Knowledge, 0);
            _.WriteEnd(StreamObjectType.Knowledge);
        });
        return new(
        [
            new CellKnowledge([new CellKnowledgeRange(other, 0, 73507), new CellKnowledgeEntry(Serial(5))]),
            new WaterlineKnowledge([new(Id(1), 73503), new(Id(2), 1, 3)]),
            new FragmentKnowledge([new(Id(3), 100000, new(0, 32768))]),
            new ContentTagKnowledge([new(Id(4), Data(4))]),
            new VersionTokenKnowledge(Data(8)),
            new UnknownKnowledge(other, unknown),
            new CellKnowledge([]),
            new WaterlineKnowledge([]),
            new FragmentKnowledge([]),
            new ContentTagKnowledge([])
        ]);
    }

    static ResponseError ChainedError() =>
        new(ResponseErrorType.Protocol, 5, "detail", new(ResponseErrorType.Win32, 32, null, new(ResponseErrorType.HResult, 0x80004005, "end")));

    static QueryChangesRequest FullQueryChanges() =>
        new(
            2,
            0xAC10,
            2,
            new(3, Cell(1)),
            3670016,
            Data(3),
            [
                new(1, 1, null, null),
                new(4, 0, new(StreamObjectType.QueryChangesFilterCellId, Data(5)), 1)
            ],
            FullKnowledge(),
            1,
            other);

    static PutChangesRequest FullPutChanges() =>
        new(4, Id(1), Id(2), 0x48, Data(3), new byte[] { 0x03, 0x00 }, other, Knowledge.Empty, 1, 0, guid);

    static IEnumerable<(string Name, object Value, Func<byte[], object> Parse)> Models()
    {
        var elements = Elements();
        for (var index = 0; index < elements.Count; index++)
        {
            yield return ($"Element{index}{elements[index].Type}", elements[index], Fsshttpb.ParseElement);
        }

        yield return ("KnowledgeEmpty", Knowledge.Empty, Fsshttpb.ParseKnowledge);
        yield return ("KnowledgeFull", FullKnowledge(), Fsshttpb.ParseKnowledge);
        yield return ("ErrorCell", ResponseError.Cell(CellErrorCode.CoherencyFailure), Fsshttpb.ParseError);
        yield return ("ErrorChained", ChainedError(), Fsshttpb.ParseError);
        yield return ("SubRequestQueryAccess", new QueryAccessRequest(1, 2, other), Fsshttpb.ParseSubRequest);
        yield return ("SubRequestQueryChanges", FullQueryChanges(), Fsshttpb.ParseSubRequest);
        yield return ("SubRequestQueryChangesMinimal", new QueryChangesRequest(3, 0), Fsshttpb.ParseSubRequest);
        yield return ("SubRequestQueryChangesNoCell", new QueryChangesRequest(3, 0x02, Arguments: new(1, null)), Fsshttpb.ParseSubRequest);
        yield return ("SubRequestPutChanges", FullPutChanges(), Fsshttpb.ParseSubRequest);
        yield return ("SubRequestPutChangesMinimal", new PutChangesRequest(5, Id(1), ExtendedGuid.Null, 0), Fsshttpb.ParseSubRequest);
        yield return ("SubRequestAllocate", new AllocateExtendedGuidRangeRequest(6, 1000), Fsshttpb.ParseSubRequest);
        yield return (
            "RequestFull",
            new FsshttpbRequest(
                new UserAgent(guid, "msword", "win", 0x39A0F666),
                [
                    new QueryAccessRequest(1, 2, other),
                    FullQueryChanges(),
                    new QueryChangesRequest(3, 0),
                    FullPutChanges(),
                    new PutChangesRequest(5, Id(1), ExtendedGuid.Null, 0),
                    new AllocateExtendedGuidRangeRequest(6, 1000, 0x4000, guid)
                ],
                new(elements),
                new(1, 0x0C),
                1,
                14),
            Fsshttpb.ParseRequest);
        yield return ("RequestMinimal", new FsshttpbRequest(null, [new QueryAccessRequest(1)]), Fsshttpb.ParseRequest);
        yield return ("RequestEmptyPackage", new FsshttpbRequest(null, [], new([])), Fsshttpb.ParseRequest);
        yield return ("RequestUserAgentGuid", new FsshttpbRequest(new(guid, null, null, 1), [new QueryAccessRequest(1)]), Fsshttpb.ParseRequest);
        yield return ("RequestUserAgentClient", new FsshttpbRequest(new(null, "client", "", 1), [new QueryAccessRequest(1)]), Fsshttpb.ParseRequest);
        yield return ("ResponseFailed", FsshttpbResponse.Failed(ResponseError.Cell(CellErrorCode.RequestNotSupported)), Fsshttpb.ParseResponse);
        yield return ("ResponseFailedChained", FsshttpbResponse.Failed(ChainedError()) with { ProtocolVersion = 14 }, Fsshttpb.ParseResponse);
        yield return (
            "ResponseFull",
            new FsshttpbResponse(
                [
                    new FailedSubResponse(1, RequestType.QueryChanges, ResponseError.Cell(CellErrorCode.CoherencyFailure)),
                    new FailedSubResponse(1, RequestType.AllocateExtendedGuidRange, ChainedError()),
                    new QueryAccessResponse(2, ResponseError.Success, ResponseError.Cell(CellErrorCode.StorageReadOnly)),
                    new QueryChangesResponse(3, Id(1), 1, FullKnowledge(), new(1, Data(20))),
                    new QueryChangesResponse(4, Id(1), 0, Knowledge.Empty),
                    new PutChangesResponse(5, Knowledge.Empty),
                    new PutChangesResponse(6, Knowledge.Empty, true),
                    new PutChangesResponse(7, FullKnowledge(), true, Id(1)),
                    new PutChangesResponse(8, Knowledge.Empty, true, ExtendedGuid.Null, [Id(2), Id(3)], 1),
                    new AllocateExtendedGuidRangeResponse(9, guid, 10, 1010)
                ],
                new(elements),
                MinimumVersion: 12),
            Fsshttpb.ParseResponse);
        yield return ("ResponseEmpty", new FsshttpbResponse([]), Fsshttpb.ParseResponse);
    }

    public static IEnumerable<string> Names() => Models().Select(_ => _.Name);

    [Test]
    [MethodDataSource(nameof(Names))]
    public async Task RoundTrip(string name)
    {
        var (_, value, parse) = Models().Single(_ => _.Name == name);
        var bytes = Fsshttpb.Write(value);
        var read = parse(bytes);
        await Assert.That(Fsshttpb.Describe(read)).IsEqualTo(Fsshttpb.Describe(value));
        await Assert.That(Fsshttpb.Hex(Fsshttpb.Write(read))).IsEqualTo(Fsshttpb.Hex(bytes));
    }

    /// <summary>
    /// The binary data of the elements and the knowledge are slices of the input, not copies.
    /// </summary>
    [Test]
    public async Task Read_Slices()
    {
        var blob = new ObjectDataBlob(Id(1), Serial(1), Data(300));
        var bytes = Fsshttpb.Write(blob);
        var read = (ObjectDataBlob) Fsshttpb.ParseElement(bytes);
        var isSlice = MemoryMarshal.TryGetArray(read.Data, out var segment) && segment.Array == bytes;
        await Assert.That(isSlice).IsTrue();
    }

    /// <summary>
    /// The Large Length of a data element larger than 32766 bytes ([MS-FSSHTTPB] 2.2.1.5.2).
    /// </summary>
    [Test]
    public async Task LargeLength()
    {
        var bytes = Fsshttpb.Write(new ObjectDataBlob(Id(1), Serial(1), Data(70000)));
        // the data element start, then the Object Data BLOB start with the length field at 32767
        var blobStart = 2 + 17 + 25 + 1;
        await Assert.That(Fsshttpb.Hex(bytes.AsSpan(blobStart, 4))).IsEqualTo("1200FEFF");
        await Assert.That(Fsshttpb.Read(bytes.AsSpan(blobStart + 4, 3).ToArray(), (ref FsshttpbReader _) => _.ReadCompactUInt64())).IsEqualTo(70000ul);
    }

    /// <summary>
    /// Models with more than one encoding are written in one: a client or platform alone with the other empty, and the
    /// data elements added of a Put Changes response after the null Storage Index when none was applied.
    /// </summary>
    [Test]
    public async Task Write_Canonical()
    {
        var request = (FsshttpbRequest) Fsshttpb.ParseRequest(new FsshttpbRequest(new(null, null, "win", 1), []).ToArray());
        await Assert.That(request.UserAgent!.Client).IsEqualTo("");
        await Assert.That(request.UserAgent.Platform).IsEqualTo("win");
        request = (FsshttpbRequest) Fsshttpb.ParseRequest(new FsshttpbRequest(new(null, "msword", null, 1), []).ToArray());
        await Assert.That(request.UserAgent!.Platform).IsEqualTo("");

        var response = (PutChangesResponse) Fsshttpb.ParseSubResponse(Fsshttpb.Write(new PutChangesResponse(1, Knowledge.Empty, DataElementsAdded: [Id(2)])));
        await Assert.That(response.HasResponseObject).IsTrue();
        await Assert.That(response.AppliedStorageIndex).IsEqualTo(ExtendedGuid.Null);
        await Assert.That(response.DataElementsAdded!.Single()).IsEqualTo(Id(2));
    }

    [Test]
    public async Task QueryChanges_Flags()
    {
        var none = new QueryChangesRequest(1, 0);
        await Assert.That((none.AllowFragments, none.IncludeFilteredOutDataElementsInKnowledge, none.RoundKnowledgeToWholeCellChanges, none.ReturnFileHash, none.UserContentEquivalentVersionOk))
            .IsEqualTo((false, false, false, false, false));
        var all = new QueryChangesRequest(1, 0x1FF, 2);
        await Assert.That((all.AllowFragments, all.IncludeFilteredOutDataElementsInKnowledge, all.RoundKnowledgeToWholeCellChanges, all.ReturnFileHash, all.UserContentEquivalentVersionOk))
            .IsEqualTo((true, true, true, true, true));
        // the I bit is in the second byte, which a one-byte field does not have
        await Assert.That(new QueryChangesRequest(1, 0x100).UserContentEquivalentVersionOk).IsFalse();
        await Assert.That(new QueryChangesRequest(1, 0x10).AllowFragments).IsTrue();
        await Assert.That(new QueryChangesArguments(0, null).IncludeStorageManifest).IsFalse();
        await Assert.That(new QueryChangesArguments(0, null).IncludeCellChanges).IsFalse();
        await Assert.That(new QueryChangesResponse(1, Id(1), 1, Knowledge.Empty).Partial).IsTrue();
        await Assert.That(new QueryChangesResponse(1, Id(1), 0, Knowledge.Empty).Partial).IsFalse();
    }

    [Test]
    public async Task PutChanges_Flags()
    {
        var none = new PutChangesRequest(1, Id(1), ExtendedGuid.Null, 0);
        await Assert.That((none.ImplyNullExpectedIfNoMapping, none.Partial, none.PartialLast, none.FavorCoherencyFailureOverNotFound, none.LastWriterWinsOnNextChange, none.AdditionalFlagBits))
            .IsEqualTo((false, false, false, false, false, (ushort) 0));
        var all = new PutChangesRequest(1, Id(1), ExtendedGuid.Null, 0xFF, AdditionalFlags: new byte[] { 0x3F, 0x00 });
        await Assert.That((all.ImplyNullExpectedIfNoMapping, all.Partial, all.PartialLast, all.FavorCoherencyFailureOverNotFound, all.LastWriterWinsOnNextChange, all.AdditionalFlagBits))
            .IsEqualTo((true, true, true, true, true, (ushort) 0x3F));
        // Additional Flags shorter than its two bytes have none
        await Assert.That(new PutChangesRequest(1, Id(1), ExtendedGuid.Null, 0, AdditionalFlags: new byte[] { 1 }).AdditionalFlagBits).IsEqualTo((ushort) 0);
    }

    [Test]
    public async Task Types()
    {
        await Assert.That(new QueryAccessRequest(1).Type).IsEqualTo(RequestType.QueryAccess);
        await Assert.That(new AllocateExtendedGuidRangeRequest(1, 1).Type).IsEqualTo(RequestType.AllocateExtendedGuidRange);
        await Assert.That(new QueryAccessResponse(1, ResponseError.Success, ResponseError.Success).Type).IsEqualTo(RequestType.QueryAccess);
        await Assert.That(new AllocateExtendedGuidRangeResponse(1, guid, 0, 1).Type).IsEqualTo(RequestType.AllocateExtendedGuidRange);
        await Assert.That(new FailedSubResponse(1, RequestType.PutChanges, ResponseError.Success).Type).IsEqualTo(RequestType.PutChanges);
        await Assert.That(new VersionTokenKnowledge(Data(1)).Kind).IsEqualTo(SpecializedKnowledge.VersionTokenKnowledgeKind);
        await Assert.That(new UnknownKnowledge(other, Data(0)).Kind).IsEqualTo(other);
        await Assert.That(ResponseError.Success).IsEqualTo(new ResponseError(ResponseErrorType.HResult, 0));
    }
}
