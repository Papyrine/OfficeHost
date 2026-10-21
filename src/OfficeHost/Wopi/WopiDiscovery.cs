/// <summary>
/// Fetches and caches the discovery of the WOPI clients: the actions they offer and their proof keys
/// ([MS-WOPI] 3.1.5.1).
/// </summary>
sealed class WopiDiscovery(IHttpClientFactory httpFactory, TimeProvider time, ILogger<WopiDiscovery> log)
{
    /// <summary>
    /// A parsed discovery.
    /// </summary>
    public record Document(IReadOnlyList<WopiAction> Actions, ProofKeys.Keys Keys);

    /// <summary>
    /// The last discovery fetched from a URL, if any, and the last attempt to fetch it, successful or not.
    /// </summary>
    sealed record Cached(Document? Document, DateTimeOffset Fetched, DateTimeOffset Attempted);

    /// <summary>
    /// A WOPI client: its gate lets one request at a time fetch its discovery, while the others read the cache.
    /// </summary>
    sealed class Client
    {
        public SemaphoreSlim Gate { get; } = new(1, 1);

        public volatile Cached? Cached;
    }

    static TimeSpan refreshInterval = TimeSpan.FromHours(1);

    /// <summary>
    /// The minimum time between two attempts to fetch a discovery that is cached, so that a stream of forged requests
    /// does not turn into a stream of discovery requests, nor a WOPI client that is down into a fetch per request.
    /// </summary>
    static TimeSpan refetchInterval = TimeSpan.FromMinutes(1);

    ConcurrentDictionary<string, Client> clients = new();

    /// <summary>
    /// The discovery, fetched when not cached or older than an hour. With <paramref name="refetch"/>, as when the keys
    /// may have been rotated, it is fetched again unless it was attempted in the last minute: then null is returned. On
    /// a failure to fetch, the previous discovery is returned, if any, and kept for a minute before the next attempt.
    /// </summary>
    public async Task<Document?> Get(string url, bool refetch = false, Cancel cancel = default)
    {
        var client = clients.GetOrAdd(url, _ => new());
        var cached = client.Cached;
        if (FromCache(cached, refetch, out var document))
        {
            return document;
        }

        await client.Gate.WaitAsync(cancel);
        try
        {
            // another request fetched it meanwhile
            var current = client.Cached;
            if (!ReferenceEquals(current, cached))
            {
                return current?.Document;
            }

            return await Fetch(url, client, cancel);
        }
        finally
        {
            client.Gate.Release();
        }
    }

    /// <summary>
    /// Whether the cache answers, without a fetch.
    /// </summary>
    bool FromCache(Cached? cached, bool refetch, out Document? document)
    {
        document = null;
        if (cached == null)
        {
            return false;
        }

        var now = time.GetUtcNow();
        var attemptedRecently = now - cached.Attempted < refetchInterval;
        if (refetch)
        {
            return attemptedRecently;
        }

        // never fetched: a WOPI client that was down may have started since
        if (cached.Document == null)
        {
            return false;
        }

        document = cached.Document;
        return attemptedRecently ||
               now - cached.Fetched < refreshInterval;
    }

    async Task<Document?> Fetch(string url, Client client, Cancel cancel)
    {
        var previous = client.Cached;
        var now = time.GetUtcNow();
        try
        {
            using var http = httpFactory.CreateClient(OfficeHostOptions.HttpClientName);
            await using var stream = await http.GetStreamAsync(url, cancel);
            var document = Parse(await XDocument.LoadAsync(stream, LoadOptions.None, cancel));
            client.Cached = new(document, now, now);
            log.LogInformation("Fetched the WOPI discovery {Url}: {Actions} actions, proof keys {Current} {Old}", url, document.Actions.Count, document.Keys.Current != null, document.Keys.Old != null);
            return document;
        }
        catch (Exception exception) when (exception is HttpRequestException or TaskCanceledException or FormatException or CryptographicException or XmlException &&
                                          !cancel.IsCancellationRequested)
        {
            log.LogError(exception, "Failed to fetch the WOPI discovery {Url}", url);
            client.Cached = new(previous?.Document, previous?.Fetched ?? now, now);
            return previous?.Document;
        }
    }

    public static Document Parse(string xml) =>
        Parse(XDocument.Parse(xml));

    static Document Parse(XDocument document)
    {
        var actions = document.Descendants()
            .Where(_ => _.Name.LocalName == "action")
            .Select(
                _ => new WopiAction(
                    (string?) _.Parent?.Attribute("name") ?? "",
                    (string?) _.Attribute("name") ?? "",
                    (string?) _.Attribute("ext") ?? "",
                    (string?) _.Attribute("urlsrc") ?? ""))
            .Where(_ => _.UrlSrc.Length > 0)
            .ToList();
        return new(actions, ProofKeys.ParseKeys(document));
    }
}
