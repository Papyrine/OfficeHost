namespace OfficeHost;

/// <summary>
/// The content of documents from SharePoint: a GET of the document's URL, signed in by the <see cref="HttpClient"/>'s
/// handler, which the application configures with an account that can read the documents (docs/shadow.md):
/// <code>
/// services.AddHttpClient&lt;IShadowContent, SharePointShadowContent&gt;()
///     .ConfigurePrimaryHttpMessageHandler(() => new HttpClientHandler { UseDefaultCredentials = true });
/// </code>
/// </summary>
public class SharePointShadowContent(HttpClient client) :
    IShadowContent
{
    public async Task<ShadowContent?> Get(ShadowDocument document, Cancel cancel = default)
    {
        using var response = await client.GetAsync(document.Url, cancel);
        if (response.StatusCode == HttpStatusCode.NotFound)
        {
            return null;
        }

        response.EnsureSuccessStatusCode();
        var content = await response.Content.ReadAsByteArrayAsync(cancel);
        return new(content, Version(response.Headers.ETag?.Tag));
    }

    /// <summary>
    /// The N of an ETag of SharePoint, <c>"{guid},N"</c>.
    /// </summary>
    internal static long? Version(string? etag)
    {
        if (etag != null &&
            long.TryParse(etag.Trim('"').Split(',')[^1], NumberStyles.None, CultureInfo.InvariantCulture, out var version))
        {
            return version;
        }

        return null;
    }
}
