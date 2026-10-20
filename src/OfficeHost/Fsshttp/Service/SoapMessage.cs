using Microsoft.AspNetCore.WebUtilities;
using Microsoft.Net.Http.Headers;

/// <summary>
/// A request to a SOAP service of Office ([MS-FSSHTTP] 2.1): the envelope, and the binary parts of an MTOM message
/// ([XOP10]) by their Content-ID.
/// </summary>
class SoapMessage(XDocument envelope, Dictionary<string, ReadOnlyMemory<byte>> parts)
{
    public static XNamespace Soap { get; } = "http://schemas.microsoft.com/sharepoint/soap/";
    public static XNamespace Envelope { get; } = "http://schemas.xmlsoap.org/soap/envelope/";
    public static XNamespace Xop { get; } = "http://www.w3.org/2004/08/xop/include";

    public XDocument Document => envelope;

    /// <summary>
    /// Reads a plain SOAP body, or an MTOM one: the first part is the envelope, the others binary parts.
    /// </summary>
    /// <exception cref="InvalidDataException">The body is not one.</exception>
    public static async Task<SoapMessage> Read(HttpRequest request, Cancel cancel)
    {
        var parts = new Dictionary<string, ReadOnlyMemory<byte>>(StringComparer.OrdinalIgnoreCase);
        try
        {
            if (!MediaTypeHeaderValue.TryParse(request.ContentType, out var type) ||
                !type.MediaType.Equals("multipart/related", StringComparison.OrdinalIgnoreCase))
            {
                return new(await XDocument.LoadAsync(request.Body, LoadOptions.None, cancel), parts);
            }

            var boundary = HeaderUtilities.RemoveQuotes(type.Boundary).Value ??
                           throw new InvalidDataException("An MTOM message without a boundary");
            var reader = new MultipartReader(boundary, request.Body)
            {
                BodyLengthLimit = null
            };
            XDocument? document = null;
            while (await reader.ReadNextSectionAsync(cancel) is { } section)
            {
                if (document == null)
                {
                    document = await XDocument.LoadAsync(section.Body, LoadOptions.None, cancel);
                    continue;
                }

                using var buffer = new MemoryStream();
                await section.Body.CopyToAsync(buffer, cancel);
                if (section.Headers != null &&
                    section.Headers.TryGetValue("Content-ID", out var id))
                {
                    parts[Normalize(id.ToString())] = buffer.ToArray();
                }
            }

            return new(document ?? throw new InvalidDataException("An MTOM message without parts"), parts);
        }
        catch (XmlException exception)
        {
            throw new InvalidDataException($"Not a SOAP message: {exception.Message}", exception);
        }
        catch (IOException exception) when (exception is not EndOfStreamException)
        {
            throw new InvalidDataException($"Not an MTOM message: {exception.Message}", exception);
        }
    }

    /// <summary>
    /// A Content-ID as in a part header (<c>&lt;id&gt;</c>) or an XOP reference (<c>cid:id</c>, URL encoded).
    /// </summary>
    static string Normalize(string id)
    {
        var value = id.Trim();
        if (value.StartsWith("cid:", StringComparison.OrdinalIgnoreCase))
        {
            value = Uri.UnescapeDataString(value[4..]);
        }

        return value.Trim('<', '>');
    }

    /// <summary>
    /// The binary data of an element: the part of its <c>xop:Include</c>, or its text as base64. Null when it has
    /// neither.
    /// </summary>
    /// <exception cref="InvalidDataException">The part is missing, or the text is not base64.</exception>
    public ReadOnlyMemory<byte>? Binary(XElement element)
    {
        var include = element.Element(Xop + "Include");
        if (include != null)
        {
            var href = (string?) include.Attribute("href") ?? "";
            if (!parts.TryGetValue(Normalize(href), out var part))
            {
                throw new InvalidDataException($"No MTOM part {href}");
            }

            return part;
        }

        var text = string.Concat(element.Nodes().OfType<XText>().Select(_ => _.Value)).Trim();
        if (text.Length == 0)
        {
            return null;
        }

        try
        {
            return Convert.FromBase64String(text);
        }
        catch (FormatException exception)
        {
            throw new InvalidDataException("Binary data that is not base64", exception);
        }
    }
}

/// <summary>
/// A response in MTOM ([MS-FSSHTTP] 2.1: the server always answers so), with binary data as parts referenced by
/// <c>xop:Include</c>.
/// </summary>
class MtomResponse
{
    const string startId = "http://tempuri.org/0";
    List<(string Id, ReadOnlyMemory<byte> Data)> parts = [];

    /// <summary>
    /// The element that references the data, written as a binary part after the envelope.
    /// </summary>
    public XElement Include(ReadOnlyMemory<byte> data)
    {
        var id = $"http://tempuri.org/1/{parts.Count}";
        parts.Add((id, data));
        return new(
            SoapMessage.Xop + "Include",
            new XAttribute("href", $"cid:{id}"),
            new XAttribute(XNamespace.Xmlns + "xop", SoapMessage.Xop));
    }

    public async Task Write(HttpResponse response, XElement envelope, Cancel cancel)
    {
        var boundary = $"urn:uuid:{Guid.NewGuid()}";
        response.ContentType =
            $"multipart/related; type=\"application/xop+xml\"; boundary=\"{boundary}\"; start=\"<{startId}>\"; start-info=\"text/xml; charset=utf-8\"";
        var body = response.Body;
        await WriteText(
            body,
            $"--{boundary}\r\nContent-ID: <{startId}>\r\nContent-Transfer-Encoding: 8bit\r\nContent-Type: application/xop+xml;charset=utf-8;type=\"text/xml; charset=utf-8\"\r\n\r\n",
            cancel);
        await WriteText(body, envelope.ToString(SaveOptions.DisableFormatting), cancel);
        foreach (var (id, data) in parts)
        {
            await WriteText(
                body,
                $"\r\n--{boundary}\r\nContent-ID: <{id}>\r\nContent-Transfer-Encoding: binary\r\nContent-Type: application/octet-stream\r\n\r\n",
                cancel);
            await body.WriteAsync(data, cancel);
        }

        await WriteText(body, $"\r\n--{boundary}--\r\n", cancel);
    }

    static async Task WriteText(Stream stream, string text, Cancel cancel) =>
        await stream.WriteAsync(Encoding.UTF8.GetBytes(text), cancel);
}
