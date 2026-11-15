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
        using var request = new HttpRequestMessage(HttpMethod.Get, document.Url);
        // SharePoint answers 403 to a client that may expect a sign-in page (no user agent of a browser) unless it says
        // it does not: then it asks for the Windows sign-in
        request.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
        request.Headers.UserAgent.ParseAdd("OfficeHost.Shadow");
        using var response = await client.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, cancel);
        if (response.StatusCode == HttpStatusCode.NotFound)
        {
            return null;
        }

        response.EnsureSuccessStatusCode();
        // a page of SharePoint's (an error, a sign-in) instead of the document
        if (string.Equals(response.Content.Headers.ContentType?.MediaType, "text/html", StringComparison.OrdinalIgnoreCase))
        {
            throw new HttpRequestException($"SharePoint answered a page, not the document: {response.RequestMessage?.RequestUri}");
        }

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
