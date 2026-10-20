using System.Text;
using Microsoft.AspNetCore.WebUtilities;
using Microsoft.Net.Http.Headers;

/// <summary>
/// Writes every request and response in full to a folder, one file each (<c>{sequence}-{method}-{path}.txt</c>), to see
/// what desktop Office sends (docs/fsshttp-checklist.md). Text bodies are written as text, others as hex, and multipart
/// (MTOM) bodies part by part. The files hold access tokens: keep them private.
/// </summary>
static class RequestLog
{
    public static void UseRequestLog(this WebApplication app, string folder)
    {
        Directory.CreateDirectory(folder);
        // after the files of a previous run, rather than over them
        var sequence = Directory.EnumerateFiles(folder).Count();
        var logger = app.Logger;
        app.Use(
            async (context, next) =>
            {
                var number = Interlocked.Increment(ref sequence);
                await Log(context, next, folder, number, logger);
            });
    }

    static async Task Log(HttpContext context, Func<Task> next, string folder, int number, ILogger logger)
    {
        var request = context.Request;
        request.EnableBuffering();
        var requestBody = await ReadAll(request.Body);
        request.Body.Position = 0;

        var response = context.Response;
        var original = response.Body;
        using var captured = new MemoryStream();
        response.Body = captured;
        try
        {
            await next();
        }
        finally
        {
            response.Body = original;
            var responseBody = captured.ToArray();
            var log = new StringBuilder();
            log.AppendLine($"#{number} {DateTimeOffset.Now:O}");
            log.AppendLine($"{request.Method} {request.Scheme}://{request.Host}{request.PathBase}{request.Path}{request.QueryString} {request.Protocol}");
            AppendHeaders(log, request.Headers);
            log.AppendLine();
            await AppendBody(log, requestBody, request.ContentType);
            log.AppendLine();
            log.AppendLine($"--- response {response.StatusCode}");
            AppendHeaders(log, response.Headers);
            log.AppendLine();
            await AppendBody(log, responseBody, response.ContentType);
            var name = $"{number:0000}-{request.Method}-{FileName(request.Path.Value)}.txt";
            await File.WriteAllTextAsync(Path.Combine(folder, name), log.ToString());
            logger.LogInformation(
                "Request #{Number} {Method} {Path} {UserAgent} -> {Status}",
                number,
                request.Method,
                request.Path.Value,
                request.Headers.UserAgent.ToString(),
                response.StatusCode);
            if (responseBody.Length > 0 && !HttpMethods.IsHead(request.Method))
            {
                await original.WriteAsync(responseBody);
            }
        }
    }

    static async Task<byte[]> ReadAll(Stream stream)
    {
        using var buffer = new MemoryStream();
        await stream.CopyToAsync(buffer);
        return buffer.ToArray();
    }

    static void AppendHeaders(StringBuilder log, IHeaderDictionary headers)
    {
        foreach (var (key, value) in headers)
        {
            log.AppendLine($"{key}: {value}");
        }
    }

    static async Task AppendBody(StringBuilder log, byte[] body, string? contentType)
    {
        if (body.Length == 0)
        {
            log.AppendLine("(no body)");
            return;
        }

        if (IsText(contentType))
        {
            log.AppendLine(Encoding.UTF8.GetString(body));
            return;
        }

        AppendHex(log, body);
        if (contentType != null &&
            contentType.Contains("multipart/", StringComparison.OrdinalIgnoreCase))
        {
            await AppendParts(log, body, contentType);
        }
    }

    static bool IsText(string? contentType)
    {
        if (contentType == null ||
            contentType.Contains("multipart/", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        string[] texts = ["text/", "xml", "json", "urlencoded", "html"];
        return texts.Any(_ => contentType.Contains(_, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>
    /// A hex dump: offset, 16 bytes, and their printable ASCII.
    /// </summary>
    static void AppendHex(StringBuilder log, byte[] body)
    {
        log.AppendLine($"({body.Length} bytes, hex)");
        for (var offset = 0; offset < body.Length; offset += 16)
        {
            var line = body.AsSpan(offset, Math.Min(16, body.Length - offset));
            log.Append($"{offset:x8}  ");
            for (var index = 0; index < 16; index++)
            {
                if (index < line.Length)
                {
                    log.Append($"{line[index]:x2} ");
                }
                else
                {
                    log.Append("   ");
                }
            }

            log.Append(' ');
            foreach (var value in line)
            {
                if (value is >= 0x20 and < 0x7f)
                {
                    log.Append((char) value);
                }
                else
                {
                    log.Append('.');
                }
            }

            log.AppendLine();
        }
    }

    /// <summary>
    /// The parts of a multipart body (MTOM/XOP), each as text or hex.
    /// </summary>
    static async Task AppendParts(StringBuilder log, byte[] body, string contentType)
    {
        try
        {
            var boundary = HeaderUtilities.RemoveQuotes(MediaTypeHeaderValue.Parse(contentType).Boundary).Value!;
            var reader = new MultipartReader(boundary, new MemoryStream(body));
            var index = 0;
            while (await reader.ReadNextSectionAsync() is { } section)
            {
                log.AppendLine($"--- part {index++}");
                foreach (var (key, value) in section.Headers ?? [])
                {
                    log.AppendLine($"{key}: {value}");
                }

                log.AppendLine();
                var content = await ReadAll(section.Body);
                if (section.ContentType != null &&
                    section.ContentType.Contains("xml", StringComparison.OrdinalIgnoreCase))
                {
                    log.AppendLine(Encoding.UTF8.GetString(content));
                }
                else
                {
                    AppendHex(log, content);
                }
            }
        }
        catch (Exception exception)
        {
            log.AppendLine($"(could not read the parts: {exception.Message})");
        }
    }

    /// <summary>
    /// The path as a file name, with the token of a desktop site as <c>~</c>.
    /// </summary>
    static string FileName(string? path)
    {
        var name = path ?? "";
        const string sites = "/fsshttp/";
        if (name.StartsWith(sites, StringComparison.OrdinalIgnoreCase))
        {
            var end = name.IndexOf('/', sites.Length);
            if (end < 0)
            {
                end = name.Length;
            }

            name = $"{sites}~{name[end..]}";
        }

        var builder = new StringBuilder();
        foreach (var value in name.Trim('/'))
        {
            if (char.IsLetterOrDigit(value) || value is '.' or '-' or '_' or '~')
            {
                builder.Append(value);
            }
            else
            {
                builder.Append('_');
            }
        }

        if (builder.Length > 80)
        {
            builder.Length = 80;
        }

        return builder.ToString();
    }
}
