using Microsoft.AspNetCore.Http.Features;

/// <summary>
/// The endpoint the module in SharePoint posts its batches to (<see cref="ShadowProtocol"/>): checks the key, reads the
/// batch, and hands it to the <see cref="ShadowInbox"/>; it never replays. 202, or 503 with Retry-After when the inbox
/// is full, so that the module backs off; 401 for a wrong key, 400 for a batch it cannot read.
/// </summary>
static class ShadowIngest
{
    public static async Task Handle(HttpContext context, ShadowInbox inbox, IOptions<OfficeHostShadowOptions> options, ILogger<ShadowInbox> log)
    {
        var request = context.Request;
        var response = context.Response;
        var settings = options.Value;
        if (!KeyMatches(request.Headers[ShadowProtocol.Key].ToString(), settings.ApiKey))
        {
            response.StatusCode = StatusCodes.Status401Unauthorized;
            return;
        }

        // a batch holds documents: beyond the default 30 MB of Kestrel
        if (context.Features.Get<IHttpMaxRequestBodySizeFeature>() is { IsReadOnly: false } limit)
        {
            limit.MaxRequestBodySize = settings.MaxBatchMegabytes * 1024L * 1024;
        }

        ShadowBatch batch;
        try
        {
            batch = await Read(request, context.RequestAborted);
        }
        catch (Exception exception) when (exception is InvalidDataException or FormatException or IOException)
        {
            log.LogWarning(new EventId((int) ShadowOutcome.ShadowFailed, nameof(ShadowOutcome.ShadowFailed)), "A batch the shadow cannot read: {Message}", exception.Message);
            response.StatusCode = StatusCodes.Status400BadRequest;
            return;
        }

        if (settings.SamplePercent is { } sample)
        {
            response.Headers[ShadowProtocol.Sample] = sample.ToString(CultureInfo.InvariantCulture);
        }

        if (!inbox.Add(batch))
        {
            response.Headers.RetryAfter = "30";
            response.StatusCode = StatusCodes.Status503ServiceUnavailable;
            return;
        }

        response.StatusCode = StatusCodes.Status202Accepted;
    }

    /// <summary>
    /// Whether the key sent is the one configured, in a time that does not tell how much of it matched. No key
    /// configured matches nothing.
    /// </summary>
    static bool KeyMatches(string sent, string configured)
    {
        if (configured.Length == 0)
        {
            return false;
        }

        return CryptographicOperations.FixedTimeEquals(SHA256.HashData(Encoding.UTF8.GetBytes(sent)), SHA256.HashData(Encoding.UTF8.GetBytes(configured)));
    }

    public static async Task<ShadowBatch> Read(HttpRequest request, Cancel cancel)
    {
        var headers = request.Headers;
        if (!Guid.TryParse(headers[ShadowProtocol.Instance], out var instance))
        {
            throw new InvalidDataException($"No {ShadowProtocol.Instance}");
        }

        var server = headers[ShadowProtocol.Server].ToString();
        var watermark = Time(headers[ShadowProtocol.Watermark].ToString(), ShadowProtocol.Watermark);
        long.TryParse(headers[ShadowProtocol.Dropped], NumberStyles.None, CultureInfo.InvariantCulture, out var dropped);
        var final = string.Equals(headers[ShadowProtocol.Final], "true", StringComparison.OrdinalIgnoreCase);
        var items = new List<ShadowItem>();
        if (request.ContentLength is not 0 &&
            MediaTypeHeaderValue.TryParse(request.ContentType, out var type) &&
            type.MediaType.Equals("multipart/mixed", StringComparison.OrdinalIgnoreCase))
        {
            var body = request.Body;
            if (string.Equals(headers.ContentEncoding, "gzip", StringComparison.OrdinalIgnoreCase))
            {
                body = new GZipStream(body, CompressionMode.Decompress);
            }

            var boundary = HeaderUtilities.RemoveQuotes(type.Boundary).Value ??
                           throw new InvalidDataException("A batch without a boundary");
            var reader = new MultipartReader(boundary, body)
            {
                BodyLengthLimit = null
            };
            while (await reader.ReadNextSectionAsync(cancel) is { } section)
            {
                // the empty section of a multipart of no parts
                if (section.Headers is not { Count: > 0 })
                {
                    continue;
                }

                items.Add(await Item(instance, server, section, cancel));
            }
        }

        return new(instance, server, watermark, dropped, final, items);
    }

    static async Task<ShadowItem> Item(Guid instance, string server, MultipartSection section, Cancel cancel)
    {
        var headers = section.Headers ?? [];
        string Header(string name) => headers.GetValueOrDefault(name).ToString();
        var sequence = long.Parse(Header(ShadowProtocol.Sequence), NumberStyles.None, CultureInfo.InvariantCulture);
        var started = Time(Header(ShadowProtocol.Started), ShadowProtocol.Started);
        if (Header(ShadowProtocol.Kind) == "tombstone")
        {
            var document = ShadowDocument.From(Header(ShadowProtocol.Document));
            return new(instance, server, sequence, started, null, document, Header(ShadowProtocol.Reason));
        }

        using var buffer = new MemoryStream();
        await section.Body.CopyToAsync(buffer, cancel);
        var bytes = buffer.ToArray();
        var length = int.Parse(Header(ShadowProtocol.RequestLength), NumberStyles.None, CultureInfo.InvariantCulture);
        if (length > bytes.Length)
        {
            throw new InvalidDataException($"{ShadowProtocol.RequestLength} {length} beyond the part of {bytes.Length} bytes");
        }

        var exchange = new Exchange(
            (int) Math.Min(sequence, int.MaxValue),
            HttpMessage.Read(bytes[..length]),
            HttpMessage.Read(bytes[length..]));
        return new(instance, server, sequence, started, exchange);
    }

    static DateTimeOffset Time(string value, string name)
    {
        if (DateTimeOffset.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal, out var time))
        {
            return time;
        }

        throw new InvalidDataException($"No {name}");
    }
}
