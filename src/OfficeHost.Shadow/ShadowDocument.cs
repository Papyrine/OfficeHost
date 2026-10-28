namespace OfficeHost;

/// <summary>
/// A document of SharePoint, by the URL desktop Office names it with: the Url of a cell storage Request, as
/// <c>http://sharepoint/sites/team/Shared%20Documents/Budget.xlsx</c>.
/// </summary>
public record ShadowDocument(Uri Url)
{
    /// <summary>
    /// The path of the document in its host, unescaped, without the leading slash: <c>sites/team/Shared Documents/Budget.xlsx</c>.
    /// </summary>
    public string Path => Uri.UnescapeDataString(Url.AbsolutePath).Trim('/');

    /// <summary>
    /// The name of the document: <c>Budget.xlsx</c>.
    /// </summary>
    public string Name => Path.Split('/')[^1];

    /// <summary>
    /// The path of its folder: <c>sites/team/Shared Documents</c>.
    /// </summary>
    public string Folder
    {
        get
        {
            var index = Path.LastIndexOf('/');
            if (index < 0)
            {
                return "";
            }

            return Path[..index];
        }
    }

    /// <summary>
    /// What tells documents apart: the host and the path, in lower case, as SharePoint does not tell cases apart.
    /// </summary>
    public string Key => $"{Url.Host}/{Path}".ToLowerInvariant();

    /// <summary>
    /// The document of a Request Url, or null when it names none.
    /// </summary>
    public static ShadowDocument? From(string? url)
    {
        if (url == null ||
            !Uri.TryCreate(url, UriKind.Absolute, out var uri))
        {
            return null;
        }

        var document = new ShadowDocument(uri);
        if (document.Name.Length == 0)
        {
            return null;
        }

        return document;
    }
}
