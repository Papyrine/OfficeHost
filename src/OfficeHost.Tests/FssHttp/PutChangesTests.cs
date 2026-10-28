/// <summary>
/// Put Changes with the options Office sets ([MS-FSSHTTPB] 2.2.2.1.4: flags 0x49, additional flags 0x8004) and without
/// them, answered as SharePoint Server SE answered the same requests (docs/fsshttp-notes.md).
/// </summary>
public class PutChangesTests
{
    /// <summary>
    /// How a case changes the Put Changes of a save, and SharePoint's answer: the error code of the SubResponse, and the
    /// Put Changes response or the Cell error of the FSSHTTPB.
    /// </summary>
    public record Case(string Name, string Answer, byte Flags = 0x49, ushort Additional = 0x8004, string Expected = "remapped", bool DropElement = false, bool ReuseId = false)
    {
        public override string ToString() => Name;
    }

    public static IEnumerable<Func<Case>> Cases() =>
    [
        () => new("save as Office", "Success applied"),
        () => new("save, no flags", "Success applied", Flags: 0, Additional: 0),
        () => new("missing element", "CellRequestFail Cell 16", DropElement: true),
        () => new("missing element, no favor coherency failure", "CellRequestFail Cell 16", Flags: 0x41, DropElement: true),
        () => new("expected not in package", "CellRequestFail Cell 16", Expected: "absent"),
        () => new("expected not in package, no favor coherency failure", "CellRequestFail Cell 16", Flags: 0x41, Expected: "absent"),
        () => new("expected null", "CellRequestFail Cell 12", Expected: "null"),
        () => new("expected null, no imply null", "Success applied", Flags: 0x48, Expected: "null"),
        () => new("expected stale", "CellRequestFail Cell 12", Expected: "stale"),
        () => new("expected stale, no imply null", "CellRequestFail Cell 12", Flags: 0x48, Expected: "stale"),
        () => new("expected without the cells", "CellRequestFail Cell 12", Expected: "partial"),
        () => new("expected without the cells, no imply null", "Success applied", Flags: 0x48, Expected: "partial"),
        () => new("id reused", "CellRequestFail Cell 112", ReuseId: true),
        () => new("id reused, no check for id reuse", "Success applied", Additional: 0, ReuseId: true)
    ];

    [Test]
    [MethodDataSource(nameof(Cases))]
    public async Task Put(Case @case)
    {
        await using var host = await TestServerHost.Start();
        var original = Zips.Create(
            ("word/document.xml", Zips.Random(3000, 1)),
            ("word/styles.xml", Zips.Random(3000, 2)),
            ("word/settings.xml", Zips.Random(3000, 3)),
            ("docProps/core.xml", Zips.Random(3000, 4)));
        await File.WriteAllBytesAsync(host.PathOf("doc.docx"), original);
        var (_, token) = await host.Open("doc.docx");
        var word = new WordClient(host, token);
        await word.Sync();
        var content = word.Content!;

        var ids = new IdAllocator(Guid.NewGuid());
        var changed = Zips.Create(
            ("word/document.xml", Zips.Random(3000, 1)),
            ("word/styles.xml", Zips.Random(3000, 2)),
            ("word/settings.xml", Zips.Random(3000, 3)),
            ("docProps/core.xml", Zips.Random(3000, 4)),
            ("customXml/item1.xml", Zips.Random(3000, 5)));
        var update = FileCell.Write(changed, content, ids, ChunkingOptions.Default);
        var index = update.Storage.StorageIndex with
        {
            Id = ids.NewId()
        };
        var package = update.Added
            .Where(_ => _ is not StorageIndex)
            .Append(index)
            .ToList();
        var expectedId = ExtendedGuid.Null;
        if (@case.Expected is "remapped" or "stale" or "partial")
        {
            var entries = content.StorageIndex.Entries.Select(_ => Remap(_, ids)).ToList();
            if (@case.Expected == "stale")
            {
                // the client expects another Storage Manifest
                entries = entries.Select(_ => _ is StorageIndexManifestMapping mapping ? mapping with { MappingId = ids.NewId() } : _).ToList();
            }

            if (@case.Expected == "partial")
            {
                // the client expects the Storage Manifest mapping alone
                entries = entries.OfType<StorageIndexManifestMapping>().Cast<StorageIndexEntry>().ToList();
            }

            var expected = content.StorageIndex with
            {
                Id = ids.NewId(),
                Entries = entries
            };
            package.Add(expected);
            expectedId = expected.Id;
        }

        if (@case.Expected == "absent")
        {
            expectedId = ids.NewId();
        }

        if (@case.DropElement)
        {
            package.Remove(package.First(_ => _ is ObjectGroup));
        }

        if (@case.ReuseId)
        {
            // the content of one Object Group under the id of another the storage has
            var groups = content.Elements.Values.OfType<ObjectGroup>().Where(_ => package.All(element => element.Id != _.Id)).Take(2).ToList();
            package.Add(groups[1] with { Id = groups[0].Id });
        }

        package = package.Select(_ => _ with { SerialNumber = SerialNumber.Null }).ToList();
        var additional = new[] { (byte) @case.Additional, (byte) (@case.Additional >> 8), (byte) 0 };
        var put = new PutChangesRequest(1, index.Id, expectedId, @case.Flags, new byte[3], additional, ClientKnowledge: word.Knowledge);
        var request = new FssHttpBRequest(new(Guid.NewGuid(), "msword", "win", 1), [put], new(package)).ToArray();
        var result = (await word.Send(WordClient.Cell("1", request, new("Coalesce", "true"), new("GetFileProps", "true"), new("BypassLockID", WordClient.SchemaLockId), new("SchemaLockID", WordClient.SchemaLockId)))).Single();
        // the properties of a put have no HaveOnlyDemotionChanges, as on SharePoint Server SE
        await Assert.That(result.Attribute("Etag")).IsNotNull();
        await Assert.That(result.Attribute("HaveOnlyDemotionChanges")).IsNull();
        var outcome = result.FssHttpB.SubResponses.Single() switch
        {
            FailedSubResponse failed => $"{failed.Error.Type} {failed.Error.Code}",
            PutChangesResponse => "applied",
            var other => other.GetType().Name
        };
        await Assert.That($"{result.ErrorCode} {outcome}").IsEqualTo(@case.Answer);

        await word.Sync();
        var file = await File.ReadAllBytesAsync(host.PathOf("doc.docx"));
        if (outcome == "applied")
        {
            await Assert.That(word.File.SameAs(file)).IsTrue();
            await Assert.That(file.SameAs(original)).IsFalse();
        }
        else
        {
            await Assert.That(file.SameAs(original)).IsTrue();
        }
    }

    /// <summary>
    /// SharePoint Server SE adds 29 bytes the spec does not define to the Put Changes response object, after the data
    /// elements added: they are read and written back as they are.
    /// </summary>
    [Test]
    public async Task Response_ExtraBytes()
    {
        var extra = Convert.FromHexString("372A09E1667B01374CBCAA1DF9E106E67A05240003054400FA28000001");
        var added = new ExtendedGuid(Guid.NewGuid(), 5);
        var response = new FssHttpBResponse([new PutChangesResponse(1, Knowledge.Empty, true, ExtendedGuid.Null, [added], Extra: extra)]);
        var read = (PutChangesResponse) FssHttpBResponse.Read(response.ToArray()).SubResponses.Single();
        await Assert.That(read.DataElementsAdded!.Single()).IsEqualTo(added);
        await Assert.That(read.Extra.ToArray().SequenceEqual(extra)).IsTrue();
    }

    static StorageIndexEntry Remap(StorageIndexEntry entry, IdAllocator ids) =>
        entry switch
        {
            StorageIndexManifestMapping value => value with { MappingSerialNumber = ids.NewSerialNumber() },
            StorageIndexCellMapping value => value with { MappingSerialNumber = ids.NewSerialNumber() },
            StorageIndexRevisionMapping value => value with { MappingSerialNumber = ids.NewSerialNumber() },
            _ => entry
        };
}
