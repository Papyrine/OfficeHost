/// <summary>
/// The SOAP envelope of a recorded body, plain or MTOM, with its binary parts by Content-ID, and the sections of an MTOM
/// body as they were, to write the body again with another envelope.
/// </summary>
class SoapBody
{
    public static XNamespace Soap { get; } = "http://schemas.microsoft.com/sharepoint/soap/";

    static XNamespace xop = "http://www.w3.org/2004/08/xop/include";

    string? boundary;
    List<(Dictionary<string, StringValues> Headers, byte[] Body)> sections = [];

    public XDocument Document { get; private set; } = null!;

    public Dictionary<string, byte[]> Parts { get; } = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// The body, or null when it is no SOAP message: a body the recorder found empty, as one of PowerPoint's
    /// (docs/replay.md).
    /// </summary>
    public static SoapBody? TryRead(HttpMessage message)
    {
        try
        {
            return Read(message);
        }
        catch (Exception exception) when (exception is IOException or InvalidDataException or XmlException)
        {
            return null;
        }
    }

    public static SoapBody Read(HttpMessage message)
    {
        var body = new SoapBody();
        if (!MediaTypeHeaderValue.TryParse(message.ContentType, out var type) ||
            !type.MediaType.Equals("multipart/related", StringComparison.OrdinalIgnoreCase))
        {
            body.Document = XDocument.Parse(Encoding.UTF8.GetString(message.Body));
            return body;
        }

        body.boundary = HeaderUtilities.RemoveQuotes(type.Boundary).Value!;
        var reader = new MultipartReader(body.boundary, new MemoryStream(message.Body))
        {
            BodyLengthLimit = null
        };
        while (reader.ReadNextSectionAsync().GetAwaiter().GetResult() is { } section)
        {
            using var buffer = new MemoryStream();
            section.Body.CopyTo(buffer);
            var headers = section.Headers ?? [];
            body.sections.Add((headers, buffer.ToArray()));
            if (body.sections.Count == 1)
            {
                body.Document = XDocument.Parse(Encoding.UTF8.GetString(buffer.ToArray()));
                continue;
            }

            if (headers.TryGetValue("Content-ID", out var id))
            {
                body.Parts[Normalize(id.ToString())] = buffer.ToArray();
            }
        }

        return body;
    }

    /// <summary>
    /// The binary data of a SubRequestData or SubResponseData: its MTOM part, or its text as base64.
    /// </summary>
    public byte[]? Binary(XElement? data)
    {
        if (data == null)
        {
            return null;
        }

        if ((string?) data.Element(xop + "Include")?.Attribute("href") is { } href)
        {
            return Parts.GetValueOrDefault(Normalize(href));
        }

        var text = string.Concat(data.Nodes().OfType<XText>().Select(_ => _.Value)).Trim();
        if (text.Length == 0)
        {
            return null;
        }

        return Convert.FromBase64String(text);
    }

    /// <summary>
    /// The body with another envelope: the other sections of an MTOM body as they were, under the same boundary.
    /// </summary>
    public byte[] With(XDocument document)
    {
        var envelope = Encoding.UTF8.GetBytes(document.Root!.ToString(SaveOptions.DisableFormatting));
        if (boundary == null)
        {
            return envelope;
        }

        using var stream = new MemoryStream();
        for (var index = 0; index < sections.Count; index++)
        {
            var (headers, body) = sections[index];
            var head = new StringBuilder($"\r\n--{boundary}\r\n");
            foreach (var (name, value) in headers)
            {
                head.Append($"{name}: {value}\r\n");
            }

            head.Append("\r\n");
            stream.Write(Encoding.UTF8.GetBytes(head.ToString()));
            if (index == 0)
            {
                stream.Write(envelope);
            }
            else
            {
                stream.Write(body);
            }
        }

        stream.Write(Encoding.UTF8.GetBytes($"\r\n--{boundary}--\r\n"));
        return stream.ToArray();
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
}
