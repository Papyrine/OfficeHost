/// <summary>
/// The file names WOPI clients send: UTF-7 encoded ([MS-WOPI] 3.3.5.1.15 and 3.3.5.1.21), and to be checked before
/// they reach the storage.
/// </summary>
static class FileNames
{
    /// <summary>
    /// The longest file name, in UTF-8 bytes: [MS-WOPI] 3.3.5.1.2.2 defaults FileNameMaxLength to 250, which also fits
    /// the 255 bytes of the common Linux file systems.
    /// </summary>
    public const int MaxBytes = 250;

    /// <summary>
    /// The characters a file name cannot hold: the path separators, and those Windows forbids, as a storage may be
    /// accessed from any OS (on NTFS, `a:b` would even name the alternate data stream `b` of the file `a`).
    /// </summary>
    static readonly SearchValues<char> forbiddenChars = SearchValues.Create("/\\<>:\"|?*");

#pragma warning disable SYSLIB0001
    public static string DecodeUtf7(string value) =>
        Encoding.UTF7.GetString(Encoding.UTF8.GetBytes(value));

    public static string EncodeUtf7(string value) =>
        Encoding.ASCII.GetString(Encoding.UTF7.GetBytes(value));
#pragma warning restore SYSLIB0001

    /// <summary>
    /// Whether the name designates a file in the current folder. Trailing dots and spaces are illegal too, as Windows
    /// drops them.
    /// </summary>
    public static bool IsLegal(string name) =>
        name.Length > 0 &&
        name is not ("." or "..") &&
        !name.AsSpan().ContainsAny(forbiddenChars) &&
        !name.Any(char.IsControl) &&
        !name.EndsWith('.') &&
        !name.EndsWith(' ') &&
        Encoding.UTF8.GetByteCount(name) <= MaxBytes;

    /// <summary>
    /// The name with the characters that would make it illegal replaced, and shortened to fit, keeping its extension.
    /// </summary>
    public static string Legal(string name)
    {
        var legal = new string(name.Select(_ => forbiddenChars.Contains(_) || char.IsControl(_) ? '_' : _).ToArray())
            .TrimEnd('.', ' ');
        if (Encoding.UTF8.GetByteCount(legal) > MaxBytes)
        {
            var (root, extension) = SplitExtension(legal);
            if (Encoding.UTF8.GetByteCount(extension) > MaxBytes / 2)
            {
                (root, extension) = (legal, "");
            }

            legal = $"{Truncate(root, MaxBytes - Encoding.UTF8.GetByteCount(extension)).TrimEnd('.', ' ')}{extension}";
        }

        if (legal.Length == 0)
        {
            return "_";
        }

        return legal;
    }

    /// <summary>
    /// The name without and with its extension: `report.docx` is (`report`, `.docx`). Leading dots do not start an
    /// extension, so `.profile` has none.
    /// </summary>
    public static (string Root, string Extension) SplitExtension(string name)
    {
        var dot = name.LastIndexOf('.');
        if (dot <= 0 ||
            name[..dot].All(_ => _ == '.'))
        {
            return (name, "");
        }

        return (name[..dot], name[dot..]);
    }

    public static string Extension(string name) =>
        SplitExtension(name).Extension;

    /// <summary>
    /// The legal name with a number before its extension: `report (2).docx`, shortening the root to fit.
    /// </summary>
    public static string Numbered(string name, int number)
    {
        var (root, extension) = SplitExtension(name);
        var suffix = $" ({number}){extension}";
        return $"{Truncate(root, MaxBytes - Encoding.UTF8.GetByteCount(suffix))}{suffix}";
    }

    /// <summary>
    /// The longest start of the value that fits in the given number of UTF-8 bytes, without splitting a character.
    /// </summary>
    static string Truncate(string value, int maxBytes)
    {
        var length = 0;
        var bytes = 0;
        foreach (var rune in value.EnumerateRunes())
        {
            bytes += rune.Utf8SequenceLength;
            if (bytes > maxBytes)
            {
                break;
            }

            length += rune.Utf16SequenceLength;
        }

        return value[..length];
    }
}
