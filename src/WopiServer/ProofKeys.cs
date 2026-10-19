/// <summary>
/// Validates the X-WOPI-Proof and X-WOPI-ProofOld headers of WOPI requests against the proof keys the WOPI clients
/// publish in their discovery ([MS-WOPI] 2.2.1 and 3.1.5.1.1.2.2.5), so that only trusted WOPI clients are served.
/// </summary>
sealed class ProofKeys(WopiDiscovery discovery, IOptions<WopiOptions> options, TimeProvider time)
{
    /// <summary>
    /// The keys of a discovery: either may be missing.
    /// </summary>
    public record Keys(RSAParameters? Current, RSAParameters? Old);

    WopiOptions Options => options.Value;

    public bool Enabled => Options.DiscoveryUrls.Count > 0;

    /// <summary>
    /// Validates the proof of a request. Returns null when the request is valid, or the reason it is not.
    /// </summary>
    /// <param name="requestUrl">The absolute URL the WOPI client called.</param>
    public async Task<string?> Validate(HttpRequest request, string accessToken, string requestUrl, Cancel cancel)
    {
        if (!Enabled)
        {
            return null;
        }

        var proof = request.Header("X-WOPI-Proof");
        if (string.IsNullOrEmpty(proof))
        {
            if (Options.RequireProof)
            {
                return "Missing X-WOPI-Proof";
            }

            return null;
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
        var proofOld = request.Header("X-WOPI-ProofOld");
        // the cached keys of all the WOPI clients first, so that the requests of one do not refetch the discovery of
        // another
        if (await Verify(expected, proof, proofOld, false, cancel) ||
            // a WOPI client may have rotated its keys since its discovery was fetched
            await Verify(expected, proof, proofOld, true, cancel))
        {
            return null;
        }

        return "The proof does not match the keys of any trusted WOPI client";
    }

    async Task<bool> Verify(byte[] expected, string proof, string? proofOld, bool refetch, Cancel cancel)
    {
        foreach (var url in Options.DiscoveryUrls)
        {
            var document = await discovery.Get(url, refetch, cancel);
            if (document != null &&
                Verify(expected, proof, proofOld, document.Keys))
            {
                return true;
            }
        }

        return false;
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

    /// <summary>
    /// The keys of the proof-key element of a discovery, from its modulus and exponent attributes, or else from the
    /// CSP blobs of its value attributes. A discovery without proof-key has no keys.
    /// </summary>
    public static Keys ParseKeys(XDocument discovery)
    {
        var element = discovery.Descendants().FirstOrDefault(_ => _.Name.LocalName == "proof-key");
        if (element == null)
        {
            return new(null, null);
        }

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
