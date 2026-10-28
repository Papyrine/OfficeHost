namespace OfficeHost;

/// <summary>
/// The files of the documents of SharePoint in the shadow's storage. The application implements it for its
/// <see cref="IOfficeStorage"/>, which has no way to create a file at a path: <see cref="FileSystemShadowDocuments"/>
/// for a <see cref="FileSystemStorage"/>, and as docs/shadow.md shows for the others. A document made by a Save As in
/// SharePoint is made by OfficeHost too, next to the file it was saved from: <see cref="Find"/> finds it there.
/// </summary>
public interface IShadowDocuments
{
    /// <summary>
    /// The id of the file of a document, or null when the storage has none.
    /// </summary>
    Task<string?> Find(ShadowDocument document, Cancel cancel = default);

    /// <summary>
    /// Adds the file of a document with its content, and returns its id.
    /// </summary>
    Task<string> Add(ShadowDocument document, Stream content, Cancel cancel = default);
}
