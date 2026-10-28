using System.Globalization;

/// <summary>
/// A request or response of a recorded exchange, as raw HTTP/1.1 (docs/replay.md): the start line, the headers, a blank
/// line, and the body as it was sent, without chunking, to the end of the file.
/// </summary>
record HttpMessage(string StartLine, List<(string Name, string Value)> Headers, byte[] Body)
{
    static byte[] separator = "\r\n\r\n"u8.ToArray();

    public string? Header(string name) =>
        Headers
            .Where(_ => _.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
            .Select(_ => _.Value)
            .FirstOrDefault();

    public string? ContentType => Header("Content-Type");

    /// <summary>
    /// The method of a request.
    /// </summary>
    public string Method => StartLine.Split(' ')[0];

    /// <summary>
    /// The path of a request, with its query.
    /// </summary>
    public string Target => StartLine.Split(' ')[1];

    /// <summary>
    /// The status of a response.
    /// </summary>
    public int Status => int.Parse(StartLine.Split(' ')[1]);

    public static HttpMessage Read(byte[] bytes)
    {
        var end = bytes.AsSpan().IndexOf(separator);
        if (end < 0)
        {
            throw new InvalidDataException("No blank line after the headers");
        }

        var lines = Encoding.UTF8.GetString(bytes, 0, end).Split("\r\n");
        var headers = lines
            .Skip(1)
            .Select(_ =>
            {
                var colon = _.IndexOf(':');
                return (_[..colon].Trim(), _[(colon + 1)..].Trim());
            })
            .ToList();
        return new(lines[0], headers, bytes[(end + separator.Length)..]);
    }

    public byte[] ToArray()
    {
        var head = new StringBuilder();
        head.Append(StartLine).Append("\r\n");
        foreach (var (name, value) in Headers)
        {
            head.Append(name).Append(": ").Append(value).Append("\r\n");
        }

        head.Append("\r\n");
        return [.. Encoding.UTF8.GetBytes(head.ToString()), .. Body];
    }
}

/// <summary>
/// A request and the response it had, numbered in the order the server received them.
/// </summary>
record Exchange(int Number, HttpMessage Request, HttpMessage Response)
{
    /// <summary>
    /// The user the recorder saw sign in, a name of the recorded server (<c>LAB\alice</c>).
    /// </summary>
    public const string UserHeader = "X-Replay-User";

    /// <summary>
    /// When the server received the request, in ISO 8601.
    /// </summary>
    public const string TimeHeader = "X-Replay-Time";

    public string User => Request.Header(UserHeader) ?? "";

    public DateTimeOffset? Time
    {
        get
        {
            if (Request.Header(TimeHeader) is { } time)
            {
                return DateTimeOffset.Parse(time, CultureInfo.InvariantCulture);
            }

            return null;
        }
    }

    /// <summary>
    /// The exchanges of a folder: <c>{number}.request.http</c> and <c>{number}.response.http</c>, by number.
    /// </summary>
    public static List<Exchange> ReadFolder(string folder) =>
        Directory.EnumerateFiles(folder, "*.request.http")
            .Select(_ =>
            {
                var number = int.Parse(Path.GetFileName(_).Split('.')[0], CultureInfo.InvariantCulture);
                var response = Path.Combine(folder, $"{Path.GetFileName(_).Split('.')[0]}.response.http");
                return new Exchange(number, HttpMessage.Read(File.ReadAllBytes(_)), HttpMessage.Read(File.ReadAllBytes(response)));
            })
            .OrderBy(_ => _.Number)
            .ToList();

    public void Write(string folder)
    {
        File.WriteAllBytes(Path.Combine(folder, $"{Number:0000}.request.http"), Request.ToArray());
        File.WriteAllBytes(Path.Combine(folder, $"{Number:0000}.response.http"), Response.ToArray());
    }
}
