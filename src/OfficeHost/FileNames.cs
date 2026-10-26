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
    static SearchValues<char> forbiddenChars = SearchValues.Create("/\\<>:\"|?*");

    /// <summary>
    /// The names Windows keeps for devices, with or without an extension.
    /// </summary>
    static FrozenSet<string> deviceNames = FrozenSet.Create(
        StringComparer.OrdinalIgnoreCase,
        [
            "CON", "PRN", "AUX", "NUL",
            "COM0", "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9", "COM¹", "COM²", "COM³",
            "LPT0", "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9", "LPT¹", "LPT²", "LPT³"
        ]);

#pragma warning disable SYSLIB0001
    /// <summary>
    /// The name of UTF-7, which is ASCII: a name that is not, from a client that sent it as is, is kept as is.
    /// </summary>
    public static string DecodeUtf7(string value)
    {
        if (!Ascii.IsValid(value))
        {
            return value;
        }

        return Encoding.UTF7.GetString(Encoding.UTF8.GetBytes(value));
    }

    public static string EncodeUtf7(string value) =>
        Encoding.ASCII.GetString(Encoding.UTF7.GetBytes(value));
#pragma warning restore SYSLIB0001

    /// <summary>
    /// Whether the name designates a file in the current folder. Trailing dots and spaces are illegal too, as Windows
    /// drops them, and so are the device names of Windows, and a half of a surrogate pair, which UTF-8 cannot hold.
    /// </summary>
    public static bool IsLegal(string name) =>
        name.Length > 0 &&
        name is not ("." or "..") &&
        !name.AsSpan().ContainsAny(forbiddenChars) &&
        !name.Any(char.IsControl) &&
        !name.EndsWith('.') &&
        !name.EndsWith(' ') &&
        !IsDeviceName(name) &&
        !HasLoneSurrogate(name) &&
        Encoding.UTF8.GetByteCount(name) <= MaxBytes;

    static bool IsDeviceName(string name)
    {
        var dot = name.IndexOf('.');
        var root = dot < 0 ? name.AsSpan() : name.AsSpan(0, dot);
        return deviceNames.GetAlternateLookup<CharSpan>().Contains(root.TrimEnd(' '));
    }

    static bool HasLoneSurrogate(CharSpan name)
    {
        for (var i = 0; i < name.Length; i++)
        {
            if (char.IsHighSurrogate(name[i]) &&
                i + 1 < name.Length &&
                char.IsLowSurrogate(name[i + 1]))
            {
                i++;
                continue;
            }

            if (char.IsSurrogate(name[i]))
            {
                return true;
            }
        }

        return false;
    }

    /// <summary>
    /// The name with the characters that would make it illegal replaced, and shortened to fit, keeping its extension.
    /// </summary>
    public static string Legal(string name)
    {
        // Replacing gives '_', never a dot or space, so trimming first trims the same.
        var length = name.AsSpan().TrimEnd(". ").Length;
        var legal = string.Create(
            length,
            name,
            static (span, name) =>
            {
                for (var i = 0; i < span.Length; i++)
                {
                    var c = name[i];
                    var pair = char.IsHighSurrogate(c) &&
                               i + 1 < span.Length &&
                               char.IsLowSurrogate(name[i + 1]);
                    if (pair)
                    {
                        span[i] = c;
                        span[i + 1] = name[i + 1];
                        i++;
                        continue;
                    }

                    if (forbiddenChars.Contains(c) ||
                        char.IsControl(c) ||
                        char.IsSurrogate(c))
                    {
                        c = '_';
                    }

                    span[i] = c;
                }
            });

        if (IsDeviceName(legal))
        {
            legal = $"_{legal}";
        }

        if (Encoding.UTF8.GetByteCount(legal) > MaxBytes)
        {
            var (root, extension) = SplitExtension(legal);
            if (Encoding.UTF8.GetByteCount(extension) > MaxBytes / 2)
            {
                root = legal;
                extension = default;
            }

            legal = $"{Truncate(root, MaxBytes - Encoding.UTF8.GetByteCount(extension)).TrimEnd(". ")}{extension}";
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
    public static SplitName SplitExtension(CharSpan name)
    {
        var dot = name.LastIndexOf('.');
        if (dot <= 0 ||
            name[..dot].TrimStart('.').IsEmpty)
        {
            return new(name, default);
        }

        return new(name[..dot], name[dot..]);
    }

    public static string Extension(string name) =>
        SplitExtension(name).Extension.ToString();

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
    static CharSpan Truncate(CharSpan value, int maxBytes)
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

/// <summary>
/// A file name split before its extension, as views of the name.
/// </summary>
readonly ref struct SplitName(CharSpan root, CharSpan extension)
{
    public CharSpan Root { get; } = root;
    public CharSpan Extension { get; } = extension;

    public void Deconstruct(out CharSpan root, out CharSpan extension)
    {
        root = Root;
        extension = Extension;
    }
}
