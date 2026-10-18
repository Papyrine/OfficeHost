return await WopiProgram.Run(args, Console.Error);

/// <summary>
/// The command line of the server: wopiserver [--config &lt;path&gt;] [--defaults &lt;path&gt;]
/// <list type="bullet">
/// <item>--config: the main configuration file, watched for changes (default /etc/wopi/wopiserver.conf)</item>
/// <item>--defaults: a system-wide defaults file, loaded over the shipped wopiserver.defaults.conf
/// (default /etc/wopi/wopiserver.defaults.conf, when present)</item>
/// </list>
/// </summary>
static class WopiProgram
{
    public const string Usage = "Usage: wopiserver [--config <path>] [--defaults <path>]";

    /// <summary>
    /// The configuration files given by the arguments, or the unknown argument.
    /// </summary>
    public static (string? DefaultsPath, string ConfigPath, string? Unknown) ParseArgs(string[] args, Func<string, bool> fileExists)
    {
        var configPath = "/etc/wopi/wopiserver.conf";
        string? defaultsPath = null;
        if (fileExists("/etc/wopi/wopiserver.defaults.conf"))
        {
            defaultsPath = "/etc/wopi/wopiserver.defaults.conf";
        }

        for (var index = 0; index < args.Length; index++)
        {
            var arg = args[index];
            if (index + 1 < args.Length &&
                arg is "--config" or "--defaults")
            {
                index++;
                if (arg == "--config")
                {
                    configPath = args[index];
                }
                else
                {
                    defaultsPath = args[index];
                }

                continue;
            }

            return (defaultsPath, configPath, arg);
        }

        return (defaultsPath, configPath, null);
    }

    /// <summary>
    /// Runs the server until it is stopped. Returns 2 for invalid arguments, 1 when the configuration is invalid.
    /// </summary>
    public static async Task<int> Run(string[] args, TextWriter error, Action<WebApplicationBuilder>? configureBuilder = null)
    {
        var (defaultsPath, configPath, unknown) = ParseArgs(args, File.Exists);
        if (unknown != null)
        {
            await error.WriteLineAsync($"Unknown argument {unknown}. {Usage}");
            return 2;
        }

        WebApplication app;
        try
        {
            app = await WopiApp.Build(
                WopiConfig.Load(defaultsPath, configPath),
                new()
                {
                    ConfigureBuilder = configureBuilder
                });
        }
        catch (Exception exception) when (exception is ConfigException or IOException or UnauthorizedAccessException or CryptographicException or FormatException or UnsupportedAppException)
        {
            await error.WriteLineAsync($"Failed to initialize the service: {exception.Message}");
            return 1;
        }

        await app.RunAsync();
        return 0;
    }
}
