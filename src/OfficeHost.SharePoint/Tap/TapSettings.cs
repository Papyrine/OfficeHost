/// <summary>
/// The settings of the module, the appSettings <c>OfficeHostShadow.*</c> of the web application that
/// Install-OfficeHostShadow.ps1 adds (docs/shadow.md, "Configuration").
/// </summary>
sealed class TapSettings
{
    /// <summary>
    /// Whether the module copies anything. Changing it changes web.config, which recycles the application pools: to
    /// pause without that, the shadow answers 503, or sets the sampling to 0.
    /// </summary>
    public bool Enabled { get; private set; }

    /// <summary>
    /// The ingest of the shadow (<c>https://shadow/officehost-shadow/exchanges</c>), or a folder to record to.
    /// </summary>
    public string? Sink { get; private set; }

    public bool IsFolder => Sink != null &&
                            !Sink.StartsWith("http://", StringComparison.OrdinalIgnoreCase) &&
                            !Sink.StartsWith("https://", StringComparison.OrdinalIgnoreCase);

    public string ApiKey { get; private set; } = "";

    /// <summary>
    /// The percentage of documents copied, by a hash of their URL: all of a document's exchanges, or none.
    /// </summary>
    public int SamplePercent { get; private set; } = 100;

    /// <summary>
    /// The most bytes of copies the module holds at once, across all requests: beyond it, exchanges become tombstones.
    /// </summary>
    public long MaxInFlightBytes { get; private set; } = 256L * 1024 * 1024;

    /// <summary>
    /// The largest exchange copied: a larger one is sent as a tombstone.
    /// </summary>
    public long MaxExchangeBytes { get; private set; } = 64L * 1024 * 1024;

    /// <summary>
    /// The path prefixes copied (<c>/sites/team</c>), none for all.
    /// </summary>
    public string[] Include { get; private set; } = [];

    public static TapSettings Read(Func<string, string?> setting)
    {
        var settings = new TapSettings
        {
            Enabled = string.Equals(setting("OfficeHostShadow.Enabled"), "true", StringComparison.OrdinalIgnoreCase),
            Sink = Trimmed(setting("OfficeHostShadow.Sink")),
            ApiKey = Trimmed(setting("OfficeHostShadow.ApiKey")) ?? ""
        };
        if (Number(setting("OfficeHostShadow.SamplePercent")) is { } sample)
        {
            settings.SamplePercent = (int) Math.Max(0, Math.Min(100, sample));
        }

        if (Number(setting("OfficeHostShadow.MaxInFlightMegabytes")) is { } inFlight)
        {
            settings.MaxInFlightBytes = inFlight * 1024 * 1024;
        }

        if (Number(setting("OfficeHostShadow.MaxExchangeMegabytes")) is { } exchange)
        {
            settings.MaxExchangeBytes = exchange * 1024 * 1024;
        }

        if (Trimmed(setting("OfficeHostShadow.Include")) is { } include)
        {
            settings.Include = include
                .Split([','], StringSplitOptions.RemoveEmptyEntries)
                .Select(_ => _.Trim())
                .Where(_ => _.Length > 0)
                .ToArray();
        }

        if (settings.Sink == null)
        {
            settings.Enabled = false;
        }

        return settings;
    }

    /// <summary>
    /// Whether a request path is copied: under one of <see cref="Include"/>, or any when there are none.
    /// </summary>
    public bool Includes(string path) =>
        Include.Length == 0 ||
        Include.Any(_ => Under(path, _));

    /// <summary>
    /// Whether a path is a prefix's or below it, by whole segments: <c>/sites/team</c> is not under <c>/sites/te</c>.
    /// </summary>
    static bool Under(string path, string prefix)
    {
        if (!path.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        return path.Length == prefix.Length ||
               prefix[prefix.Length - 1] == '/' ||
               path[prefix.Length] == '/';
    }

    static string? Trimmed(string? value)
    {
        var trimmed = value?.Trim();
        if (string.IsNullOrEmpty(trimmed))
        {
            return null;
        }

        return trimmed;
    }

    static long? Number(string? value)
    {
        if (long.TryParse(value?.Trim(), NumberStyles.None, CultureInfo.InvariantCulture, out var number))
        {
            return number;
        }

        return null;
    }
}
