using System.Buffers.Binary;
using System.Xml.Linq;

/// <summary>
/// Validates the X-WOPI-Proof and X-WOPI-ProofOld headers of WOPI requests against the proof keys the WOPI clients
/// publish in their discovery ([MS-WOPI] 2.2.1 and 3.1.5.1.1.2.2.5), so the server knows a request comes from a trusted
/// WOPI client. The Python implementation has no such validation.
/// </summary>
class ProofKeys(WopiConfig config, HttpClient http, TimeProvider time, ILogger log)
{
    /// <summary>
    /// The keys of a discovery: either may be missing.
    /// </summary>
    public record Keys(RSAParameters? Current, RSAParameters? Old);

    record CachedKeys(Keys Keys, DateTimeOffset Fetched);

    static readonly TimeSpan refreshInterval = TimeSpan.FromHours(1);

    /// <summary>
    /// The minimum age of the cached keys before a mismatch triggers a refetch, so a stream of forged requests does not
    /// turn into a stream of discovery requests.
    /// </summary>
    static readonly TimeSpan rotationRefetchInterval = TimeSpan.FromMinutes(1);

    readonly SemaphoreSlim gate = new(1, 1);
    readonly Dictionary<string, CachedKeys> cache = [];

    public IReadOnlyList<string> DiscoveryUrls =>
        config.Get("security", "proofkeydiscoveryurls", "").Split((char[]?) null, StringSplitOptions.RemoveEmptyEntries);

    public bool Enabled => DiscoveryUrls.Count > 0;

    /// <summary>
    /// Whether requests without proof are refused. The requests of the bridge to this server are not signed.
    /// </summary>
    public bool Required => Enabled && config.IsTrue("security", "proofkeyrequired", false);

    /// <summary>
    /// Validates the proof of a request. Returns null when the request is valid, or the reason it is not.
    /// </summary>
    /// <param name="requestUrl">The absolute URL the WOPI client called.</param>
    public async Task<string?> Validate(HttpRequest request, string accessToken, string requestUrl)
    {
        var proof = request.Header("X-WOPI-Proof");
        var proofOld = request.Header("X-WOPI-ProofOld");
        if (string.IsNullOrEmpty(proof))
        {
            return "Missing X-WOPI-Proof";
        }

        if (!long.TryParse(request.Header("X-WOPI-TimeStamp"), CultureInfo.InvariantCulture, out var ticks) ||
            ticks is < 0 or > 0x2BCA2875F4373FFF)
        {
            return "Missing or invalid X-WOPI-TimeStamp";
        }

        var age = time.GetUtcNow().UtcDateTime - new DateTime(ticks, DateTimeKind.Utc);
        if (age > TimeSpan.FromMinutes(20) ||
            age < TimeSpan.FromMinutes(-5))
        {
            return "X-WOPI-TimeStamp is outside of the accepted window";
        }

        var expected = ExpectedProof(accessToken, requestUrl, ticks);
        foreach (var url in DiscoveryUrls)
        {
            var keys = await GetKeys(url, false);
            if (keys != null &&
                Verify(expected, proof, proofOld, keys))
            {
                return null;
            }

            // the WOPI client may have rotated its keys since we fetched them
            keys = await GetKeys(url, true);
            if (keys != null &&
                Verify(expected, proof, proofOld, keys))
            {
                return null;
            }
        }

        return "The proof does not match the keys of any trusted WOPI client";
    }

    /// <summary>
    /// The signed data: the token, the uppercased request URL and the timestamp, each prefixed with its length as a
    /// 4-byte big-endian integer. The timestamp is an 8-byte big-endian integer.
    /// </summary>
    public static byte[] ExpectedProof(string accessToken, string requestUrl, long ticks)
    {
        var token = Encoding.UTF8.GetBytes(accessToken);
        var url = Encoding.UTF8.GetBytes(requestUrl.ToUpperInvariant());
        var expected = new byte[4 + token.Length + 4 + url.Length + 4 + 8];
        var span = expected.AsSpan();
        BinaryPrimitives.WriteInt32BigEndian(span, token.Length);
        token.CopyTo(span[4..]);
        span = span[(4 + token.Length)..];
        BinaryPrimitives.WriteInt32BigEndian(span, url.Length);
        url.CopyTo(span[4..]);
        span = span[(4 + url.Length)..];
        BinaryPrimitives.WriteInt32BigEndian(span, 8);
        BinaryPrimitives.WriteInt64BigEndian(span[4..], ticks);
        return expected;
    }

    /// <summary>
    /// A request is valid when X-WOPI-Proof matches the current or the old key, or X-WOPI-ProofOld matches the current
    /// key: the latter is the case of a WOPI client that rotated its keys after the discovery was fetched.
    /// </summary>
    static bool Verify(byte[] expected, string proof, string? proofOld, Keys keys) =>
        Verify(expected, proof, keys.Current) ||
        Verify(expected, proof, keys.Old) ||
        Verify(expected, proofOld, keys.Current);

    static bool Verify(byte[] expected, string? signature, RSAParameters? key)
    {
        if (string.IsNullOrEmpty(signature) ||
            key == null)
        {
            return false;
        }

        try
        {
            using var rsa = RSA.Create(key.Value);
            return rsa.VerifyData(expected, Convert.FromBase64String(signature), HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
        }
        catch (Exception exception) when (exception is FormatException or CryptographicException)
        {
            return false;
        }
    }

    async Task<Keys?> GetKeys(string url, bool refetch)
    {
        await gate.WaitAsync();
        try
        {
            var now = time.GetUtcNow();
            if (cache.TryGetValue(url, out var cached))
            {
                var age = now - cached.Fetched;
                if (refetch ? age < rotationRefetchInterval : age < refreshInterval)
                {
                    if (refetch)
                    {
                        // fetched too recently to be worth another attempt
                        return null;
                    }

                    return cached.Keys;
                }
            }

            try
            {
                var keys = ParseDiscovery(await http.GetStringAsync(url));
                cache[url] = new(keys, now);
                log.LogInformation("Fetched WOPI proof keys {url} {current} {old}", url, keys.Current != null, keys.Old != null);
                return keys;
            }
            catch (Exception exception) when (exception is HttpRequestException or TaskCanceledException or FormatException or CryptographicException or System.Xml.XmlException)
            {
                log.LogError("Failed to fetch WOPI proof keys {url} {error}", url, exception.Message);
                // keep using the previous keys, if any
                return cached?.Keys;
            }
        }
        finally
        {
            gate.Release();
        }
    }

    /// <summary>
    /// The keys of the proof-key element of a discovery XML, from its modulus and exponent attributes, or else from the
    /// CSP blobs of its value attributes.
    /// </summary>
    public static Keys ParseDiscovery(string xml)
    {
        var element = XDocument.Parse(xml).Descendants().FirstOrDefault(_ => _.Name.LocalName == "proof-key") ??
                      throw new FormatException("The discovery has no proof-key element");
        return new(
            Key(element, "modulus", "exponent", "value"),
            Key(element, "oldmodulus", "oldexponent", "oldvalue"));
    }

    static RSAParameters? Key(XElement element, string modulus, string exponent, string blob)
    {
        var modulusValue = (string?) element.Attribute(modulus);
        var exponentValue = (string?) element.Attribute(exponent);
        if (!string.IsNullOrEmpty(modulusValue) &&
            !string.IsNullOrEmpty(exponentValue))
        {
            return new RSAParameters
            {
                Modulus = Convert.FromBase64String(modulusValue),
                Exponent = Convert.FromBase64String(exponentValue)
            };
        }

        var blobValue = (string?) element.Attribute(blob);
        if (string.IsNullOrEmpty(blobValue))
        {
            return null;
        }

        using var rsa = new RSACryptoServiceProvider();
        rsa.ImportCspBlob(Convert.FromBase64String(blobValue));
        return rsa.ExportParameters(false);
    }
}
