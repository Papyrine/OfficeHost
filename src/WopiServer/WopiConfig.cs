/// <summary>
/// The INI configuration of the server (the same `wopiserver.conf` format as the Python implementation),
/// read with Python <c>configparser</c> semantics. Values are read live, so edits to the reloadable
/// configuration file apply without a restart.
/// </summary>
public class WopiConfig(IConfiguration configuration)
{
    public IConfiguration Configuration { get; } = configuration;

    /// <summary>
    /// The value, or null when the option is missing.
    /// </summary>
    public string? Get(string section, string key) =>
        Configuration[$"{section}:{key}"];

    public string Get(string section, string key, string fallback) =>
        Get(section, key) ?? fallback;

    /// <summary>
    /// The value, failing with a <see cref="ConfigException"/> when the option is missing.
    /// </summary>
    public string GetRequired(string section, string key) =>
        Get(section, key) ?? throw new ConfigException($"No option '{key}' in section: '{section}'");

    public int GetInt(string section, string key) =>
        ParseInt(section, key, GetRequired(section, key));

    public int GetInt(string section, string key, int fallback)
    {
        var value = Get(section, key);
        if (value == null)
        {
            return fallback;
        }

        return ParseInt(section, key, value);
    }

    static int ParseInt(string section, string key, string value)
    {
        if (int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var result))
        {
            return result;
        }

        throw new ConfigException($"Invalid integer '{value}' for option '{key}' in section: '{section}'");
    }

    public bool Has(string section, string key) =>
        Get(section, key) != null;

    /// <summary>
    /// <c>config.get(section, key, fallback=fallback).upper() == 'TRUE'</c>.
    /// </summary>
    public bool IsTrue(string section, string key, bool fallback) =>
        string.Equals(Get(section, key, fallback.ToString()), "TRUE", StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// <c>config.get(section, key, fallback=fallback).upper() in ('TRUE', 'YES')</c>.
    /// </summary>
    public bool IsTrueOrYes(string section, string key, bool fallback)
    {
        var value = Get(section, key, fallback.ToString());
        return string.Equals(value, "TRUE", StringComparison.OrdinalIgnoreCase) ||
               string.Equals(value, "YES", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// The WOPI access token validity, in seconds.
    /// </summary>
    public int TokenValidity => GetInt("general", "tokenvalidity", 86400);

    /// <summary>
    /// The WOPI lock expiration, in seconds. Microsoft requires WOPI locks to expire after 30 minutes.
    /// </summary>
    public int LockExpiration => GetInt("general", "wopilockexpiration", 1800);

    public string LogLevelName => Get("general", "loglevel", "Info");

    public bool IsDebug => LogLevelName == "Debug";

    public LogLevel LogLevel =>
        LogLevelName switch
        {
            "Critical" => LogLevel.Critical,
            "Error" => LogLevel.Error,
            "Warning" => LogLevel.Warning,
            "Debug" => LogLevel.Debug,
            _ => LogLevel.Information
        };

    /// <summary>
    /// Loads the configuration: the shipped defaults, then the optional system-wide defaults, then the main
    /// configuration file, which is watched for changes.
    /// </summary>
    public static IConfigurationRoot Load(string? defaultsPath, string configPath)
    {
        var builder = new ConfigurationBuilder()
            .AddIniFile(Path.Combine(AppContext.BaseDirectory, "wopiserver.defaults.conf"), optional: true, reloadOnChange: false);
        if (defaultsPath != null)
        {
            builder.AddIniFile(defaultsPath, optional: false, reloadOnChange: false);
        }

        return builder
            .AddIniFile(configPath, optional: true, reloadOnChange: true)
            .Build();
    }
}

/// <summary>
/// A missing or invalid configuration option: this is fatal at startup.
/// </summary>
public class ConfigException(string message) :
    Exception(message);
