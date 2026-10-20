/// <summary>
/// The validation of X-WOPI-Proof against the proof keys of the WOPI client discovery.
/// </summary>
public class ProofKeyTests
{
    const string discoveryUrl = "https://office/hosting/discovery";

    /// <summary>
    /// A WOPI client: it publishes its keys in a discovery, and signs its requests.
    /// </summary>
    sealed class WopiClient : IDisposable
    {
        public RSA Current { get; set; } = RSA.Create(2048);
        public RSA? Old { get; set; } = RSA.Create(2048);
        public bool CspBlobsOnly { get; init; }
        public bool Unreachable { get; set; }
        public string Url { get; init; } = discoveryUrl;
        public int Fetches { get; private set; }

        /// <summary>
        /// The fetches, and the attempts while unreachable.
        /// </summary>
        public int Attempts { get; private set; }

        public Task<HttpResponseMessage> Handle(HttpRequestMessage request)
        {
            Attempts++;
            if (request.RequestUri!.ToString() != Url ||
                Unreachable)
            {
                throw new HttpRequestException($"Unreachable {request.RequestUri}");
            }

            Fetches++;
            return Task.FromResult(
                new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new StringContent(Discovery())
                });
        }

        string Discovery()
        {
            var attributes = $"{Attributes("", Current)} {(Old == null ? "" : Attributes("old", Old))}";
            return $"""
                    <?xml version="1.0" encoding="utf-8"?>
                    <wopi-discovery>
                      <net-zone name="external-https">
                        <app name="Word" favIconUrl="https://office/word.ico" checkLicense="true">
                          <action name="edit" ext="docx" urlsrc="https://office/we/wordeditorframe.aspx?" />
                        </app>
                      </net-zone>
                      <proof-key {attributes} />
                    </wopi-discovery>
                    """;
        }

        string Attributes(string prefix, RSA key)
        {
            var parameters = key.ExportParameters(false);
            using var csp = new RSACryptoServiceProvider();
            csp.ImportParameters(parameters);
            var blob = $"{prefix}value=\"{Convert.ToBase64String(csp.ExportCspBlob(false))}\"";
            if (CspBlobsOnly)
            {
                return blob;
            }

            return $"{blob} {prefix}modulus=\"{Convert.ToBase64String(parameters.Modulus!)}\" {prefix}exponent=\"{Convert.ToBase64String(parameters.Exponent!)}\"";
        }

        public void Dispose()
        {
            Current.Dispose();
            Old?.Dispose();
        }
    }

    static string Sign(RSA key, string token, string url, long ticks) =>
        Convert.ToBase64String(key.SignData(ProofKeys.ExpectedProof(token, url, ticks), HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1));

    sealed class Context(TestServerHost host, WopiClient client, ManualTime time, string fileId, string token) : IAsyncDisposable
    {
        public TestServerHost Host { get; } = host;
        public WopiClient Client { get; } = client;
        public ManualTime Time { get; } = time;
        public string FileId { get; } = fileId;
        public string Token { get; } = token;

        public string Url => $"http://localhost/wopi/files/{FileId}?access_token={Token}";

        public long Now => Time.Now.UtcTicks;

        /// <summary>
        /// A CheckFileInfo with the given proof headers.
        /// </summary>
        public async Task<HttpResponseMessage> CheckFileInfo(string? proof, string? proofOld = null, long? ticks = null)
        {
            using var request = new HttpRequestMessage(HttpMethod.Get, Url);
            request.Headers.Add("X-WOPI-TimeStamp", (ticks ?? Now).ToString());
            if (proof != null)
            {
                request.Headers.Add("X-WOPI-Proof", proof);
            }

            if (proofOld != null)
            {
                request.Headers.Add("X-WOPI-ProofOld", proofOld);
            }

            return await Host.Client.SendAsync(request);
        }

        public Task<HttpResponseMessage> CheckFileInfo(RSA key, long? ticks = null) =>
            CheckFileInfo(Sign(key, Token, Url, ticks ?? Now), ticks: ticks);

        public async ValueTask DisposeAsync()
        {
            await Host.DisposeAsync();
            Client.Dispose();
        }
    }

    /// <param name="other">Another trusted WOPI client, listed before the one of the context.</param>
    static async Task<Context> Start(bool required = false, bool cspBlobsOnly = false, WopiClient? other = null)
    {
        var time = new ManualTime();
        var client = new WopiClient
        {
            CspBlobsOnly = cspBlobsOnly
        };
        var host = await TestServerHost.Start(
            _ =>
            {
                if (other != null)
                {
                    _.DiscoveryUrls.Add(other.Url);
                }

                _.DiscoveryUrls.Add(discoveryUrl);
                _.RequireProof = required;
            },
            time);
        host.Http = request =>
        {
            if (other != null &&
                request.RequestUri!.ToString() == other.Url)
            {
                return other.Handle(request);
            }

            return client.Handle(request);
        };
        await File.WriteAllTextAsync(host.PathOf("/test.docx"), "content");
        var (fileId, token) = await host.Open("/test.docx");
        return new(host, client, time, fileId, token);
    }

    [Test]
    public async Task ExpectedProof_Layout()
    {
        var expected = ProofKeys.ExpectedProof("ab", "http://x/y?z=é", 0x0102030405060708);
        byte[] url = [.. "HTTP://X/Y?Z=É"u8];
        byte[] layout =
        [
            0, 0, 0, 2, (byte) 'a', (byte) 'b',
            0, 0, 0, (byte) url.Length, .. url,
            0, 0, 0, 8, 1, 2, 3, 4, 5, 6, 7, 8
        ];
        await Assert.That(expected.SequenceEqual(layout)).IsTrue();
    }

    [Test]
    public async Task ParseDiscovery_ModulusAndExponentOrCspBlobs()
    {
        foreach (var cspBlobsOnly in new[] { false, true })
        {
            using var client = new WopiClient
            {
                CspBlobsOnly = cspBlobsOnly
            };
            using var response = await client.Handle(new(HttpMethod.Get, discoveryUrl));
            var keys = WopiDiscovery.Parse(await response.Content.ReadAsStringAsync()).Keys;
            await Assert.That(keys.Current!.Value.Modulus!.SequenceEqual(client.Current.ExportParameters(false).Modulus!)).IsTrue();
            await Assert.That(keys.Old!.Value.Modulus!.SequenceEqual(client.Old!.ExportParameters(false).Modulus!)).IsTrue();
        }
    }

    [Test]
    public async Task SignedWithCurrentKey()
    {
        await using var context = await Start();
        using var response = await context.CheckFileInfo(context.Client.Current);
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.OK);
    }

    [Test]
    public async Task SignedWithCspBlobKeys()
    {
        await using var context = await Start(cspBlobsOnly: true);
        using var response = await context.CheckFileInfo(context.Client.Current);
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.OK);
    }

    /// <summary>
    /// The WOPI client rotated its keys and the host fetched the new discovery: X-WOPI-Proof may still be signed with
    /// the old key.
    /// </summary>
    [Test]
    public async Task SignedWithOldKey()
    {
        await using var context = await Start();
        using var response = await context.CheckFileInfo(context.Client.Old!);
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.OK);
    }

    /// <summary>
    /// The WOPI client rotated its keys since the host fetched the discovery: X-WOPI-ProofOld is signed with the key the
    /// host knows as current.
    /// </summary>
    [Test]
    public async Task ProofOldSignedWithCurrentKey()
    {
        await using var context = await Start();
        using var newKey = RSA.Create(2048);
        var ticks = context.Now;
        using var response = await context.CheckFileInfo(
            Sign(newKey, context.Token, context.Url, ticks),
            Sign(context.Client.Current, context.Token, context.Url, ticks),
            ticks);
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.OK);
    }

    [Test]
    public async Task SignedWithAnotherKey()
    {
        await using var context = await Start();
        using var attacker = RSA.Create(2048);
        using var response = await context.CheckFileInfo(attacker);
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.InternalServerError);
        await Assert.That(response.Headers.Contains("X-WOPI-ServerError")).IsTrue();
    }

    /// <summary>
    /// A proof is only valid for the request it was made for: another URL, token or timestamp.
    /// </summary>
    [Test]
    public async Task SignedForAnotherRequest()
    {
        await using var context = await Start();
        var key = context.Client.Current;
        var ticks = context.Now;
        foreach (var proof in new[]
                 {
                     Sign(key, context.Token, context.Url.Replace("/files/", "/files/x"), ticks),
                     Sign(key, $"{context.Token}x", context.Url, ticks),
                     Sign(key, context.Token, context.Url, ticks - 1)
                 })
        {
            using var response = await context.CheckFileInfo(proof, ticks: ticks);
            await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.InternalServerError);
        }

        using (var garbage = await context.CheckFileInfo("not base64!"))
        {
            await Assert.That(garbage.StatusCode).IsEqualTo(HttpStatusCode.InternalServerError);
        }
    }

    /// <summary>
    /// The proof is checked with a timestamp at most 20 minutes old, and not too far in the future.
    /// </summary>
    [Test]
    [Arguments(-25)]
    [Arguments(10)]
    public async Task TimestampOutsideWindow(int minutes)
    {
        await using var context = await Start();
        using var response = await context.CheckFileInfo(context.Client.Current, context.Time.Now.AddMinutes(minutes).UtcTicks);
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.InternalServerError);
    }

    [Test]
    public async Task MissingTimestamp()
    {
        await using var context = await Start();
        using var request = new HttpRequestMessage(HttpMethod.Get, context.Url);
        request.Headers.Add("X-WOPI-Proof", Sign(context.Client.Current, context.Token, context.Url, context.Now));
        using var response = await context.Host.Client.SendAsync(request);
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.InternalServerError);
    }

    /// <summary>
    /// Unsigned requests are accepted unless RequireProof is set: Collabora Online does not sign its requests.
    /// </summary>
    [Test]
    [Arguments(false, HttpStatusCode.OK)]
    [Arguments(true, HttpStatusCode.InternalServerError)]
    public async Task Unsigned(bool required, HttpStatusCode expected)
    {
        await using var context = await Start(required);
        using var response = await context.Host.Client.GetAsync(context.Url);
        await Assert.That(response.StatusCode).IsEqualTo(expected);
    }

    /// <summary>
    /// The POST operations and GetFile/PutFile are validated too.
    /// </summary>
    [Test]
    public async Task OtherOperations()
    {
        await using var context = await Start();
        using var attacker = RSA.Create(2048);
        foreach (var (key, expected) in new[] { (context.Client.Current, HttpStatusCode.OK), (attacker, HttpStatusCode.InternalServerError) })
        {
            var lockUrl = context.Url;
            using (var request = new HttpRequestMessage(HttpMethod.Post, lockUrl))
            {
                request.Headers.Add("X-WOPI-Override", "GET_LOCK");
                request.Headers.Add("X-WOPI-TimeStamp", context.Now.ToString());
                request.Headers.Add("X-WOPI-Proof", Sign(key, context.Token, lockUrl, context.Now));
                using var response = await context.Host.Client.SendAsync(request);
                await Assert.That(response.StatusCode).IsEqualTo(expected);
            }

            var contentsUrl = $"http://localhost/wopi/files/{context.FileId}/contents?access_token={context.Token}";
            using (var request = new HttpRequestMessage(HttpMethod.Get, contentsUrl))
            {
                request.Headers.Add("X-WOPI-TimeStamp", context.Now.ToString());
                request.Headers.Add("X-WOPI-Proof", Sign(key, context.Token, contentsUrl, context.Now));
                using var response = await context.Host.Client.SendAsync(request);
                await Assert.That(response.StatusCode).IsEqualTo(expected);
            }
        }
    }

    /// <summary>
    /// The WOPI client rotates its keys: once the cached keys are over a minute old, a mismatch refetches the
    /// discovery. Forged requests do not turn into a discovery request each.
    /// </summary>
    [Test]
    public async Task KeyRotation()
    {
        await using var context = await Start();
        using (var response = await context.CheckFileInfo(context.Client.Current))
        {
            await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.OK);
        }

        await Assert.That(context.Client.Fetches).IsEqualTo(1);
        using var attacker = RSA.Create(2048);
        for (var attempt = 0; attempt < 3; attempt++)
        {
            using var forged = await context.CheckFileInfo(attacker);
            await Assert.That(forged.StatusCode).IsEqualTo(HttpStatusCode.InternalServerError);
        }

        await Assert.That(context.Client.Fetches).IsEqualTo(1);

        // the client rotates: the current key becomes the old one
        context.Client.Old!.Dispose();
        context.Client.Old = context.Client.Current;
        context.Client.Current = RSA.Create(2048);
        context.Time.Advance(TimeSpan.FromMinutes(2));
        using (var response = await context.CheckFileInfo(context.Client.Current))
        {
            await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.OK);
        }

        await Assert.That(context.Client.Fetches).IsEqualTo(2);
    }

    /// <summary>
    /// Signed requests cannot be validated while the discovery is unreachable; unsigned ones are not affected.
    /// </summary>
    [Test]
    public async Task DiscoveryUnreachable()
    {
        await using var context = await Start();
        context.Client.Unreachable = true;
        using (var signed = await context.CheckFileInfo(context.Client.Current))
        {
            await Assert.That(signed.StatusCode).IsEqualTo(HttpStatusCode.InternalServerError);
        }

        using var unsigned = await context.Host.Client.GetAsync(context.Url);
        await Assert.That(unsigned.StatusCode).IsEqualTo(HttpStatusCode.OK);
    }

    /// <summary>
    /// Keys fetched before the discovery became unreachable keep being used, and the discovery is attempted again once a
    /// minute rather than on each request.
    /// </summary>
    [Test]
    public async Task DiscoveryUnreachable_KeepsPreviousKeys()
    {
        await using var context = await Start();
        using (var first = await context.CheckFileInfo(context.Client.Current))
        {
            await Assert.That(first.StatusCode).IsEqualTo(HttpStatusCode.OK);
        }

        context.Client.Unreachable = true;
        context.Time.Advance(TimeSpan.FromHours(2));
        for (var attempt = 0; attempt < 3; attempt++)
        {
            using var response = await context.CheckFileInfo(context.Client.Current);
            await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.OK);
        }

        await Assert.That(context.Client.Attempts).IsEqualTo(2);
        context.Time.Advance(TimeSpan.FromMinutes(2));
        using (var later = await context.CheckFileInfo(context.Client.Current))
        {
            await Assert.That(later.StatusCode).IsEqualTo(HttpStatusCode.OK);
        }

        await Assert.That(context.Client.Attempts).IsEqualTo(3);
    }

    /// <summary>
    /// A request is checked against the cached keys of all the WOPI clients before a discovery is refetched, so that
    /// the requests of one client do not refetch the discovery of another.
    /// </summary>
    [Test]
    public async Task SeveralClients()
    {
        using var other = new WopiClient
        {
            Url = "https://other/hosting/discovery"
        };
        await using var context = await Start(other: other);
        for (var attempt = 0; attempt < 3; attempt++)
        {
            using var response = await context.CheckFileInfo(context.Client.Current);
            await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.OK);
            context.Time.Advance(TimeSpan.FromMinutes(2));
        }

        await Assert.That(other.Fetches).IsEqualTo(1);
        await Assert.That(context.Client.Fetches).IsEqualTo(1);
    }
}
