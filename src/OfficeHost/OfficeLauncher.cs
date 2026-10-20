namespace OfficeHost;

/// <summary>
/// Opens files in a WOPI client: issues the access token and builds the URL of the action from the discovery of the
/// client ([MS-WOPI] 3.1.5.1.1.2.3). The page of the application then POSTs the token to the URL, as an
/// <c>access_token</c> and an <c>access_token_ttl</c> form field, usually in an iframe.
/// </summary>
public sealed partial class OfficeLauncher
{
    IOfficeStorage storage;
    IOfficeTokens tokens;
    IOptions<OfficeHostOptions> options;
    WopiDiscovery discovery;

    internal OfficeLauncher(IOfficeStorage storage, IOfficeTokens tokens, IOptions<OfficeHostOptions> options, WopiDiscovery discovery)
    {
        this.storage = storage;
        this.tokens = tokens;
        this.options = options;
        this.discovery = discovery;
    }

    [GeneratedRegex("<[^>]*>")]
    private static partial Regex Placeholder();

    /// <summary>
    /// Launches the given action, as edit or view, on the file of the access.
    /// </summary>
    /// <param name="request">The current request, for the URL of the application when <see cref="OfficeHostOptions.BaseUrl"/> is not set.</param>
    /// <param name="discoveryUrl">The WOPI client to use, one of <see cref="OfficeHostOptions.DiscoveryUrls"/>, or null for the first that offers the action.</param>
    /// <exception cref="InvalidOperationException">No WOPI client offers the action for the type of the file: a <see cref="OfficeClientUnreachableException"/> when the discovery of some could not be fetched.</exception>
    public async Task<OfficeLaunch> Launch(OfficeAccess access, string action, HttpRequest? request = null, string? discoveryUrl = null, Cancel cancel = default)
    {
        var settings = options.Value;
        var file = await storage.GetFile(access.FileId, cancel) ??
                   throw new FileNotFoundException($"No file {access.FileId}");
        var extension = FileNames.Extension(file.Name).TrimStart('.');
        var urlSrc = await UrlSrc(settings, discoveryUrl, action, extension, cancel);
        string baseUrl;
        if (settings.BaseUrl != null)
        {
            baseUrl = settings.BaseUrl.TrimEnd('/');
        }
        else if (request != null)
        {
            baseUrl = request.BaseUrl(settings);
        }
        else
        {
            throw new InvalidOperationException($"{nameof(OfficeHostOptions)}.{nameof(OfficeHostOptions.BaseUrl)} is needed to launch outside of a request");
        }

        var token = await tokens.Issue(access, cancel);
        // the placeholders like <ui=UI_LLCC&> are optional: drop them
        var url = Placeholder().Replace(urlSrc, "");
        var separator = "&";
        if (url.EndsWith('?') ||
            url.EndsWith('&'))
        {
            separator = "";
        }
        else if (!url.Contains('?'))
        {
            separator = "?";
        }

        url = $"{url}{separator}WOPISrc={Uri.EscapeDataString(WopiHandler.FileUrl(baseUrl, file.Id))}";
        return new(url, token.Value, token.Expires.ToUnixTimeMilliseconds());
    }

    async Task<string> UrlSrc(OfficeHostOptions settings, string? discoveryUrl, string action, string extension, Cancel cancel)
    {
        IEnumerable<string> urls = settings.DiscoveryUrls;
        if (discoveryUrl != null)
        {
            urls = [discoveryUrl];
        }

        var unreachable = new List<string>();
        foreach (var url in urls)
        {
            var document = await discovery.Get(url, cancel: cancel);
            if (document == null)
            {
                unreachable.Add(url);
                continue;
            }

            var match = document.Actions.FirstOrDefault(
                _ => _.Name == action &&
                     string.Equals(_.Extension, extension, StringComparison.OrdinalIgnoreCase));
            if (match != null)
            {
                return match.UrlSrc;
            }
        }

        var message = $"No WOPI client offers the {action} action for .{extension} files";
        if (unreachable.Count > 0)
        {
            throw new OfficeClientUnreachableException($"{message}, and the discovery of {string.Join(", ", unreachable)} could not be fetched", unreachable);
        }

        throw new InvalidOperationException(message);
    }
}