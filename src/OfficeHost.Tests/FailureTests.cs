/// <summary>
/// The handling of storage failures, through a <see cref="FaultyStorage"/>: a missing file is a 404, any other
/// failure a 500 with X-WOPI-ServerError, and the file is left as it was.
/// </summary>
public class FailureTests
{
    const string databuf = "file content";

    static string? Header(HttpResponseMessage response, string name)
    {
        if (response.Headers.TryGetValues(name, out var values))
        {
            return string.Join(", ", values);
        }

        return null;
    }

    sealed class Context(TestServerHost host, FaultyStorage storage, string fileId, string token) : IAsyncDisposable
    {
        public TestServerHost Host { get; } = host;
        public FaultyStorage Storage { get; } = storage;
        public string FileId { get; } = fileId;
        public string Token { get; } = token;

        public void Fail(string operation, Exception exception) =>
            Storage.Fault = _ => _ == operation ? exception : null;

        public Task<HttpResponseMessage> Post(string operation, params (string Name, string Value)[] headers) =>
            Host.Post(FileId, Token, operation, headers);

        public ValueTask DisposeAsync() =>
            Host.DisposeAsync();
    }

    static async Task<Context> Start(string content = databuf)
    {
        FaultyStorage? storage = null;
        var host = await TestServerHost.Start(wrapStorage: _ => storage = new(_));
        await File.WriteAllTextAsync(host.PathOf("/test.docx"), content);
        var (fileId, token) = await host.Open("/test.docx");
        return new(host, storage!, fileId, token);
    }

    static async Task IsServerError(HttpResponseMessage response)
    {
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.InternalServerError);
        await Assert.That(string.IsNullOrEmpty(Header(response, "X-WOPI-ServerError"))).IsFalse();
    }

    [Test]
    public async Task GetFile_ReadFails()
    {
        await using var context = await Start();
        context.Fail(nameof(IOfficeStorage.OpenRead), new IOException("disk"));
        using var response = await context.Host.Client.GetAsync($"/wopi/files/{context.FileId}/contents?access_token={context.Token}");
        await IsServerError(response);
    }

    [Test]
    public async Task PutFile_WriteFails()
    {
        await using var context = await Start();
        using (var locked = await context.Post("LOCK", ("X-WOPI-Lock", "L")))
        {
            await Assert.That(locked.StatusCode).IsEqualTo(HttpStatusCode.OK);
        }

        context.Fail(nameof(IOfficeStorage.Write), new IOException("disk"));
        using var response = await context.Host.PutFile(context.FileId, context.Token, "new", ("X-WOPI-Lock", "L"));
        await IsServerError(response);
        await Assert.That(await File.ReadAllTextAsync(context.Host.PathOf("/test.docx"))).IsEqualTo(databuf);
    }

    /// <summary>
    /// A file deleted between the checks of PutFile and the write.
    /// </summary>
    [Test]
    public async Task PutFile_FileDeletedBeforeWrite()
    {
        await using var context = await Start("");
        context.Fail(nameof(IOfficeStorage.Write), new FileNotFoundException("gone"));
        using var response = await context.Host.PutFile(context.FileId, context.Token, "new");
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.NotFound);
    }

    [Test]
    public async Task PutRelativeFile_CreateFails()
    {
        await using var context = await Start();
        context.Fail(nameof(IOfficeStorage.CreateSibling), new IOException("disk"));
        using var response = await context.Post("PUT_RELATIVE", ("X-WOPI-SuggestedTarget", ".pdf"));
        await IsServerError(response);
    }

    /// <summary>
    /// Finding a free name needs the storage to tell whether a name is taken.
    /// </summary>
    [Test]
    public async Task PutRelativeFile_GetSiblingFails()
    {
        await using var context = await Start();
        context.Fail(nameof(IOfficeStorage.GetSibling), new IOException("disk"));
        using var response = await context.Post("PUT_RELATIVE", ("X-WOPI-SuggestedTarget", ".pdf"));
        await IsServerError(response);
        await Assert.That(File.Exists(context.Host.PathOf("/test.pdf"))).IsFalse();
    }

    /// <summary>
    /// A storage that does not allow creating files in the folder: the WOPI client is told Save As is not supported.
    /// </summary>
    [Test]
    public async Task PutRelativeFile_AccessDenied()
    {
        await using var context = await Start();
        context.Fail(nameof(IOfficeStorage.CreateSibling), new UnauthorizedAccessException("read only folder"));
        using var response = await context.Post("PUT_RELATIVE", ("X-WOPI-SuggestedTarget", ".pdf"));
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.NotImplemented);
    }

    [Test]
    public async Task DeleteFile_StorageFails()
    {
        await using var context = await Start();
        context.Fail(nameof(IOfficeStorage.Delete), new IOException("disk"));
        using var response = await context.Post("DELETE");
        await IsServerError(response);
        await Assert.That(File.Exists(context.Host.PathOf("/test.docx"))).IsTrue();
    }

    [Test]
    public async Task RenameFile_StorageFails()
    {
        await using var context = await Start();
        context.Fail(nameof(IOfficeStorage.Rename), new IOException("disk"));
        using var response = await context.Post("RENAME_FILE", ("X-WOPI-RequestedName", "renamed"));
        await IsServerError(response);
        await Assert.That(File.Exists(context.Host.PathOf("/test.docx"))).IsTrue();
    }

    [Test]
    public async Task RenameFile_FileDeletedBeforeRename()
    {
        await using var context = await Start();
        context.Fail(nameof(IOfficeStorage.Rename), new FileNotFoundException("gone"));
        using var response = await context.Post("RENAME_FILE", ("X-WOPI-RequestedName", "renamed"));
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.NotFound);
    }

    /// <summary>
    /// A failure the endpoints do not expect, here of the storage in CheckFileInfo, is a logged server error.
    /// </summary>
    [Test]
    [Arguments(typeof(IOException))]
    [Arguments(typeof(InvalidOperationException))]
    public async Task UnexpectedException(Type type)
    {
        await using var context = await Start();
        context.Fail(nameof(IOfficeStorage.GetFile), (Exception) Activator.CreateInstance(type, "unexpected")!);
        using var response = await context.Host.Client.GetAsync($"/wopi/files/{context.FileId}?access_token={context.Token}");
        await IsServerError(response);
        await Assert.That(Header(response, "X-WOPI-ServerVersion")).IsNotNull();
        await Assert.That(context.Host.Log.Lines.Any(_ => _.StartsWith("Error") && _.Contains("Unexpected failure"))).IsTrue();
    }
}
