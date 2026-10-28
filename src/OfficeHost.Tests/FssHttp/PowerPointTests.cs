/// <summary>
/// Desktop PowerPoint, as it went against SharePoint Server and the sample (docs/fsshttp-notes.md): WhoAmI and
/// GetDocMetaInfo first, as Excel, then a coauthoring session, as Word, with an exclusive lock id to fall back to; the
/// Cell queries of the editors table and of the content, without the metadata partition; IsOnlyClient polls; saves; and
/// an exit on close.
/// </summary>
public class PowerPointTests
{
    static byte[] Deck(string text) => Zips.Create(("ppt/slides/slide1.xml", Zips.Text(text)), ("ppt/media/image1.png", Zips.Random(20_000, 5)));

    [Test]
    public async Task EditAndSave()
    {
        await using var host = await TestServerHost.Start();
        await File.WriteAllBytesAsync(host.PathOf("deck.pptx"), Deck("start"));
        var (fileId, token) = await host.Open("deck.pptx");
        var powerPoint = new WordClient(host, token);

        var metadata = await powerPoint.Send(WordClient.SubRequest("WhoAmI", "1"), WordClient.SubRequest("GetDocMetaInfo", "2", null, "1"));
        await Assert.That(metadata.Select(_ => _.ErrorCode).ToList()).IsEquivalentTo(["Success", "Success"]);

        // the open: the join, with the exclusive lock PowerPoint would take if coauthoring were not supported
        var editors = new XAttribute("PartitionID", "7808F4DD-2385-49D6-B7CE-37ACA5E43602");
        var open = await powerPoint.Send(
            WordClient.SubRequest("Coauth", "1", WordClient.Data(powerPoint.CoauthAttributes("JoinCoauthoring"), new XAttribute("Timeout", 2580), new XAttribute("AllowFallbackToExclusive", "true"), new XAttribute("ExclusiveLockID", "{736FD68C-0FF5-4F65-B5F5-8356F7DA996A}"))),
            WordClient.Cell("7", WordClient.Query(), editors),
            WordClient.Cell("6", WordClient.Query(), new XAttribute("GetFileProps", "true")),
            WordClient.SubRequest("GetVersions", "9"),
            WordClient.SubRequest("ServerTime", "11"));
        var byToken = open.ToDictionary(_ => _.Token);
        await Assert.That(byToken["1"].Attribute("LockType")).IsEqualTo("SchemaLock");
        await Assert.That(byToken["1"].Attribute("CoauthStatus")).IsEqualTo("Alone");
        await Assert.That(byToken["7"].ErrorCode).IsEqualTo("Success");
        await Assert.That(byToken["6"].ErrorCode).IsEqualTo("Success");
        await Assert.That(byToken["9"].ErrorCode).IsEqualTo("RequestNotSupported");
        await Assert.That((await host.Service<IOfficeLockStore>().Get(fileId))!.Id).IsEqualTo($"fsshttp-schema:{WordClient.SchemaLockId}");
        await Assert.That(await powerPoint.IsOnlyClient()).IsTrue();

        // the save, and the close
        await powerPoint.Sync();
        await Assert.That(await powerPoint.Save(Deck("saved"))).IsTypeOf<PutChangesResponse>();
        await Assert.That((await File.ReadAllBytesAsync(host.PathOf("deck.pptx"))).SameAs(Deck("saved"))).IsTrue();
        await Assert.That((await powerPoint.Coauth("ExitCoauthoring")).ErrorCode).IsEqualTo("Success");
        await Assert.That(await host.Service<IOfficeLockStore>().Get(fileId)).IsNull();
    }
}
