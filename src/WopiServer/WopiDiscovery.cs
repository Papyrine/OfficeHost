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

    record Cached(Document Document, DateTimeOffset Fetched);

    static TimeSpan refreshInterval = TimeSpan.FromHours(1);

    /// <summary>
    /// The minimum age of a cached discovery before a refetch is allowed, so that a stream of forged requests does not
    /// turn into a stream of discovery requests.
    /// </summary>
    static TimeSpan refetchInterval = TimeSpan.FromMinutes(1);

    SemaphoreSlim gate = new(1, 1);
    Dictionary<string, Cached> cache = [];

    /// <summary>
    /// The discovery, fetched when not cached or older than an hour. With <paramref name="refetch"/>, as when the keys
    /// may have been rotated, it is fetched again unless it was in the last minute: then null is returned. On a failure
    /// to fetch, the previous discovery is returned, if any.
    /// </summary>
    public async Task<Document?> Get(string url, bool refetch = false, Cancel cancel = default)
    {
        await gate.WaitAsync(cancel);
        try
        {
            var now = time.GetUtcNow();
            if (cache.TryGetValue(url, out var cached))
            {
                var age = now - cached.Fetched;
                if (refetch &&
                    age < refetchInterval)
                {
                    return null;
                }

                if (!refetch &&
                    age < refreshInterval)
                {
                    return cached.Document;
                }
            }

            try
            {
                using var http = httpFactory.CreateClient(WopiOptions.HttpClientName);
                var document = Parse(await http.GetStringAsync(url, cancel));
                cache[url] = new(document, now);
                log.LogInformation("Fetched the WOPI discovery {Url}: {Actions} actions, proof keys {Current} {Old}", url, document.Actions.Count, document.Keys.Current != null, document.Keys.Old != null);
                return document;
            }
            catch (Exception exception) when (exception is HttpRequestException or TaskCanceledException or FormatException or CryptographicException or XmlException)
            {
                log.LogError(exception, "Failed to fetch the WOPI discovery {Url}", url);
                return cached?.Document;
            }
        }
        finally
        {
            gate.Release();
        }
    }

    public static Document Parse(string xml)
    {
        var document = XDocument.Parse(xml);
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