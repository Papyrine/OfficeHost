/// <summary>
/// Python <c>urllib.parse</c> quoting. The WOPISrc and the bridge URLs are built with these exact rules,
/// which differ from <see cref="WebUtility.UrlEncode"/> (that one leaves `!*()` alone and encodes `~`).
/// </summary>
static class UrlCodec
{
    /// <summary>
    /// <c>quote_plus</c>: every byte but `A-Za-z0-9_.-~` is percent-encoded, and spaces become `+`.
    /// </summary>
    public static string QuotePlus(string value)
    {
        var builder = new StringBuilder(value.Length);
        foreach (var b in Encoding.UTF8.GetBytes(value))
        {
            if (b == ' ')
            {
                builder.Append('+');
            }
            else if (IsUnreserved(b))
            {
                builder.Append((char) b);
            }
            else
            {
                builder.Append('%').Append(b.ToString("X2", CultureInfo.InvariantCulture));
            }
        }

        return builder.ToString();
    }

    static bool IsUnreserved(byte b) =>
        b is
            >= (byte) 'A' and <= (byte) 'Z' or
            >= (byte) 'a' and <= (byte) 'z' or
            >= (byte) '0' and <= (byte) '9' or
            (byte) '_' or
            (byte) '.' or
            (byte) '-' or
            (byte) '~';

    /// <summary>
    /// <c>unquote_plus</c>: `+` becomes a space, then percent escapes are decoded as UTF-8.
    /// </summary>
    public static string UnquotePlus(string value) =>
        Unquote(value.Replace('+', ' '));

    /// <summary>
    /// <c>unquote</c>: percent escapes are decoded as UTF-8, invalid sequences are left as they are.
    /// </summary>
    public static string Unquote(string value)
    {
        if (!value.Contains('%'))
        {
            return value;
        }

        var bytes = new List<byte>(value.Length);
        var builder = new StringBuilder(value.Length);
        for (var index = 0; index < value.Length; index++)
        {
            if (value[index] == '%' &&
                index + 2 < value.Length &&
                char.IsAsciiHexDigit(value[index + 1]) &&
                char.IsAsciiHexDigit(value[index + 2]))
            {
                bytes.Add(byte.Parse(value.AsSpan(index + 1, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture));
                index += 2;
                continue;
            }

            Flush(bytes, builder);
            builder.Append(value[index]);
        }

        Flush(bytes, builder);
        return builder.ToString();
    }

    static void Flush(List<byte> bytes, StringBuilder builder)
    {
        if (bytes.Count == 0)
        {
            return;
        }

        builder.Append(Encoding.UTF8.GetString(bytes.ToArray()));
        bytes.Clear();
    }

    /// <summary>
    /// <c>urlencode</c> of a sequence of pairs, with values quoted by <see cref="QuotePlus"/>.
    /// </summary>
    public static string UrlEncode(IEnumerable<KeyValuePair<string, string>> parameters) =>
        string.Join('&', parameters.Select(_ => $"{QuotePlus(_.Key)}={QuotePlus(_.Value)}"));
}
