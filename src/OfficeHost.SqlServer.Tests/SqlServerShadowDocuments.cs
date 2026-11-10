/// <summary>
/// The documents of SharePoint in a <see cref="SqlServerStorage"/>: each in the folder of its host and path
/// (<c>sharepoint/sites/team/Shared Documents</c>), by its name. A document a Save As makes, which OfficeHost makes next
/// to the file it was saved from, is found there too.
/// </summary>
// begin-snippet: SqlServerShadowDocuments
class SqlServerShadowDocuments(SqlServerStorage storage) :
    IShadowDocuments
{
    public async Task<string?> Find(ShadowDocument document, Cancel cancel) =>
        (await storage.GetFile(FolderOf(document), document.Name, cancel))?.Id;

    public async Task<string> Add(ShadowDocument document, Stream content, Cancel cancel)
    {
        // throws when the folder has a file of that name
        var file = await storage.Add(FolderOf(document), document.Name, "shadow", content, cancel);
        return file.Id;
    }

    static string FolderOf(ShadowDocument document) =>
        $"{document.Url.Host}/{document.Folder}";
}
// end-snippet
