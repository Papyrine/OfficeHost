/// <summary>
/// Small Python string idioms the WOPI server relies on, mostly for logging.
/// </summary>
static partial class PyStrings
{
    /// <summary>
    /// The last 20 characters (Python <c>s[-20:]</c>), used to log tokens without disclosing them.
    /// </summary>
    public static string Short(this string? value)
    {
        if (value == null)
        {
            return "";
        }

        if (value.Length <= 20)
        {
            return value;
        }

        return value[^20..];
    }

    /// <summary>
    /// Python <c>str.title()</c>: `REFRESH_LOCK` becomes `Refresh_Lock`, `CheckFileInfo` becomes `Checkfileinfo`.
    /// </summary>
    public static string Title(this string value)
    {
        var builder = new StringBuilder(value.Length);
        var previousIsCased = false;
        foreach (var c in value)
        {
            if (previousIsCased)
            {
                builder.Append(char.ToLowerInvariant(c));
            }
            else
            {
                builder.Append(char.ToUpperInvariant(c));
            }

            previousIsCased = char.IsLetter(c);
        }

        return builder.ToString();
    }

    /// <summary>
    /// <c>value.encode().decode('utf-7')</c>: WOPI clients send the target names of PutRelative and
    /// RenameFile UTF-7 encoded.
    /// </summary>
#pragma warning disable SYSLIB0001
    public static string DecodeUtf7(string value) =>
        Encoding.UTF7.GetString(Encoding.UTF8.GetBytes(value));
#pragma warning restore SYSLIB0001

    static readonly HashSet<string> windowsDeviceFiles =
    [
        "CON", "PRN", "AUX", "NUL",
        "COM0", "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9",
        "LPT0", "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9"
    ];

    [GeneratedRegex("[^A-Za-z0-9_.-]")]
    private static partial Regex UnsafeFilenameChars();

    /// <summary>
    /// Werkzeug <c>secure_filename</c>: an ASCII-only file name safe to store on the local file system.
    /// </summary>
    public static string SecureFilename(string filename)
    {
        var normalized = filename.Normalize(NormalizationForm.FormKD);
        var ascii = new string(normalized.Where(char.IsAscii).ToArray());
        ascii = ascii.Replace('/', ' ');
        if (OperatingSystem.IsWindows())
        {
            ascii = ascii.Replace('\\', ' ');
        }

        var joined = string.Join('_', ascii.Split((char[]?) null, StringSplitOptions.RemoveEmptyEntries));
        var result = UnsafeFilenameChars().Replace(joined, "").Trim('.', '_');
        if (OperatingSystem.IsWindows() &&
            result.Length > 0 &&
            windowsDeviceFiles.Contains(result.Split('.')[0].ToUpperInvariant()))
        {
            result = $"_{result}";
        }

        return result;
    }
}
