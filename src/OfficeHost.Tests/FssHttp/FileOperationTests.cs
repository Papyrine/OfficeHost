using System.Xml.Linq;

/// <summary>
/// FileOperation Rename, the rename of Office's title bar, answered as SharePoint Server SE answers it
/// (src/sharepoint-lab/Compare-CellStorage.ps1, docs/fsshttp-notes.md).
/// </summary>
public class FileOperationTests
{
    const string exclusiveLockId = "{6E5F3C2A-1B4D-4E8F-9A7C-2D1E0F3B4A5C}";

    static async Task<(TestServerHost Host, string FileId, string Token, WordClient Word)> Start(Mode mode = Mode.Edit)
    {
        var host = await TestServerHost.Start();
        await File.WriteAllBytesAsync(host.PathOf("doc.docx"), Zips.Create(("word/document.xml", Zips.Text("rename me"))));
        await File.WriteAllBytesAsync(host.PathOf("other.docx"), Zips.Create(("word/document.xml", Zips.Text("other"))));
        var (fileId, token) = await host.Open("doc.docx", mode);
        return (host, fileId, token, new(host, token));
    }

    static async Task<SubResult> Rename(WordClient client, string? newName, string? lockId = null, string? operation = "Rename")
    {
        var attributes = new List<XAttribute>();
        if (operation != null) { attributes.Add(new("FileOperation", operation)); }
        if (newName != null) { attributes.Add(new("NewFileName", newName)); }
        if (lockId != null) { attributes.Add(new("ExclusiveLockID", lockId)); }
        return (await client.Send(WordClient.SubRequest("FileOperation", "1", WordClient.Data(attributes.ToArray())))).Single();
    }

    /// <summary>
    /// The file is renamed, and its site keeps working under the new name: the token in its path stands for the file.
    /// </summary>
    [Test]
    public async Task Rename()
    {
        var (host, _, token, word) = await Start();
        await using var _ = host;
        await word.Sync();
        var renamed = await Rename(word, "renamed.docx");
        await Assert.That(renamed.ErrorCode).IsEqualTo("Success");
        await Assert.That(File.Exists(host.PathOf("renamed.docx"))).IsTrue();
        await Assert.That(File.Exists(host.PathOf("doc.docx"))).IsFalse();

        var bytes = await host.Client.GetByteArrayAsync($"/fsshttp/{token}/renamed.docx");
        await Assert.That(bytes.SameAs(await File.ReadAllBytesAsync(host.PathOf("renamed.docx")))).IsTrue();
        await word.Sync();
        await Assert.That(word.File.SameAs(bytes)).IsTrue();

        // to another extension too, as SharePoint allows; and to the same name
        await Assert.That((await Rename(word, "renamed.txt")).ErrorCode).IsEqualTo("Success");
        await Assert.That((await Rename(word, "renamed.txt")).ErrorCode).IsEqualTo("Success");
        await Assert.That(File.Exists(host.PathOf("renamed.txt"))).IsTrue();
    }

    /// <summary>
    /// A file in a coauthoring session, or under a lock other than the exclusive lock of the ExclusiveLockID, is not
    /// renamed; the answer names no one, as SharePoint's.
    /// </summary>
    [Test]
    public async Task Locks()
    {
        var (host, fileId, token, word) = await Start();
        await using var _ = host;

        await word.Send(word.Join());
        var coauthoring = await Rename(word, "renamed.docx");
        await Assert.That(coauthoring.ErrorCode).IsEqualTo("FileAlreadyLockedOnServer");
        await Assert.That(coauthoring.ErrorMessage).IsNull();
        await word.Coauth("ExitCoauthoring");

        await word.Send(WordClient.SubRequest("ExclusiveLock", "1", WordClient.Data(new XAttribute("ExclusiveLockRequestType", "GetLock"), new XAttribute("ExclusiveLockID", exclusiveLockId), new XAttribute("Timeout", 600))));
        await Assert.That((await Rename(word, "renamed.docx")).ErrorCode).IsEqualTo("FileAlreadyLockedOnServer");
        await Assert.That((await Rename(word, "renamed.docx", "{00000000-0000-0000-0000-000000000001}")).ErrorCode).IsEqualTo("FileAlreadyLockedOnServer");
        await Assert.That((await Rename(word, "renamed.docx", exclusiveLockId)).ErrorCode).IsEqualTo("Success");
        await word.Send(WordClient.SubRequest("ExclusiveLock", "1", WordClient.Data(new XAttribute("ExclusiveLockRequestType", "ReleaseLock"), new XAttribute("ExclusiveLockID", exclusiveLockId))));

        using (var wopi = await host.Post(fileId, token, "LOCK", ("X-WOPI-Lock", "wopi-lock")))
        {
            await Assert.That(wopi.StatusCode).IsEqualTo(HttpStatusCode.OK);
        }

        await Assert.That((await Rename(word, "again.docx")).ErrorCode).IsEqualTo("FileAlreadyLockedOnServer");
        await Assert.That(File.Exists(host.PathOf("renamed.docx"))).IsTrue();
    }

    [Test]
    public async Task Errors()
    {
        var (host, _, _, word) = await Start();
        await using var _ = host;
        await Assert.That((await Rename(word, "doc?.docx")).ErrorCode).IsEqualTo("InvalidArgument");
        await Assert.That((await Rename(word, "other.docx")).ErrorCode).IsEqualTo("FileAlreadyExists");
        await Assert.That((await Rename(word, "renamed.docx", operation: null)).ErrorCode).IsEqualTo("InvalidArgument");
        // no name renames nothing
        await Assert.That((await Rename(word, null)).ErrorCode).IsEqualTo("Success");
        await Assert.That(File.Exists(host.PathOf("doc.docx"))).IsTrue();
    }

    [Test]
    public async Task WithoutRenameAccess()
    {
        var (host, _, _, word) = await Start(Mode.ReadOnly);
        await using var _ = host;
        await Assert.That((await Rename(word, "renamed.docx")).ErrorCode).IsEqualTo("FileUnauthorizedAccess");
        await Assert.That(File.Exists(host.PathOf("doc.docx"))).IsTrue();
    }

    [Test]
    public async Task MissingFile()
    {
        var (host, _, _, word) = await Start();
        await using var _ = host;
        File.Delete(host.PathOf("doc.docx"));
        await Assert.That((await Rename(word, "renamed.docx")).ErrorCode).IsEqualTo("FileNotExistsOrCannotBeCreated");
    }
}
