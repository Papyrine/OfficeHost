using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Net;

/// <summary>
/// Sends batches to the ingest of the shadow (docs/shadow.md, "Protocol"), with the TLS and proxy settings of the
/// process: nothing of <see cref="ServicePointManager"/> is changed, which would change SharePoint's own calls.
/// </summary>
sealed class HttpSink(TapSettings settings, Guid instance, string server) : ITapSink
{
    public TapSent Send(IReadOnlyList<TapItem> items, DateTime watermark, long dropped, bool final)
    {
        var request = (HttpWebRequest) WebRequest.Create(settings.Sink!);
        request.Method = "POST";
        request.Timeout = 30_000;
        request.ReadWriteTimeout = 30_000;
        request.Headers["X-OfficeHost-Shadow-Key"] = settings.ApiKey;
        request.Headers["X-Shadow-Instance"] = instance.ToString();
        request.Headers["X-Shadow-Server"] = server;
        request.Headers["X-Shadow-Watermark"] = watermark.ToString("o", CultureInfo.InvariantCulture);
        request.Headers["X-Shadow-Dropped"] = dropped.ToString(CultureInfo.InvariantCulture);
        if (final)
        {
            request.Headers["X-Shadow-Final"] = "true";
        }

        if (items.Count == 0)
        {
            request.ContentLength = 0;
        }
        else
        {
            var boundary = Guid.NewGuid().ToString("N");
            request.ContentType = $"multipart/mixed; boundary=\"{boundary}\"";
            request.Headers["Content-Encoding"] = "gzip";
            // gzip written to the request as it goes, chunked: neither a copy of the batch in memory, nor one that
            // HttpWebRequest keeps to send it again (the ingest asks for no authentication to send it again for)
            request.SendChunked = true;
            request.AllowWriteStreamBuffering = false;
            using var stream = request.GetRequestStream();
            TapBatch.Write(items, boundary, stream);
        }

        try
        {
            using var response = (HttpWebResponse) request.GetResponse();
            return new()
            {
                Sent = response.StatusCode == HttpStatusCode.Accepted,
                SamplePercent = Sample(response),
                Error = $"HTTP {(int) response.StatusCode}"
            };
        }
        catch (WebException exception) when (exception.Response is HttpWebResponse response)
        {
            using (response)
            {
                var sent = new TapSent
                {
                    SamplePercent = Sample(response),
                    Error = $"HTTP {(int) response.StatusCode}"
                };
                if (response.StatusCode == HttpStatusCode.ServiceUnavailable)
                {
                    sent.RetryAfter = TimeSpan.FromSeconds(30);
                    if (int.TryParse(response.Headers["Retry-After"], NumberStyles.None, CultureInfo.InvariantCulture, out var seconds))
                    {
                        sent.RetryAfter = TimeSpan.FromSeconds(Math.Min(seconds, 600));
                    }
                }

                return sent;
            }
        }
        catch (Exception exception) when (exception is WebException or IOException)
        {
            return new()
            {
                Error = exception.Message
            };
        }
    }

    static int? Sample(HttpWebResponse response)
    {
        if (int.TryParse(response.Headers["X-Shadow-Sample"], NumberStyles.None, CultureInfo.InvariantCulture, out var sample))
        {
            return sample;
        }

        return null;
    }
}
