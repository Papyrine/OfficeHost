namespace OfficeHost;

/// <summary>
/// The content of documents as SharePoint has them, for the shadow to start a document from when the traffic does not
/// hold it all (Office opened it from its cache). <see cref="SharePointShadowContent"/> gets it from SharePoint.
/// </summary>
public interface IShadowContent
{
    /// <summary>
    /// The content of a document now, or null when SharePoint has none.
    /// </summary>
    Task<ShadowContent?> Get(ShadowDocument document, Cancel cancel = default);
}

/// <param name="Version">The version of the file in SharePoint, the N of its ETag (<c>"{guid},N"</c>), to check it
/// against the traffic.</param>
public record ShadowContent(byte[] Content, long? Version);
