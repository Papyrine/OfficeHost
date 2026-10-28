/// <summary>
/// The map of the keys of SharePoint's documents to the ids of their files, as an application keeps it in its own
/// database.
/// </summary>
interface IShadowDocumentMap
{
    Task<string?> Get(string key, Cancel cancel);

    Task Set(string key, string fileId, Cancel cancel);
}

// begin-snippet: SqlServerShadowDocuments
class SqlServerShadowDocuments(SqlServerStorage storage, IShadowDocumentMap map) :
    IShadowDocuments
{
    public Task<string?> Find(ShadowDocument document, Cancel cancel) =>
        map.Get(document.Key, cancel);

    public async Task<string> Add(ShadowDocument document, Stream content, Cancel cancel)
    {
        var folder = $"{document.Url.Host}/{document.Folder}";
        var file = await storage.Add(folder, document.Name, "shadow", content, cancel);
        await map.Set(document.Key, file!.Id, cancel);
        return file.Id;
    }
}
// end-snippet
