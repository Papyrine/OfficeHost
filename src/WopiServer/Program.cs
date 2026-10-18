// The WOPI server. Usage: wopiserver [--config <path>] [--defaults <path>]
//   --config    the main configuration file, watched for changes (default /etc/wopi/wopiserver.conf)
//   --defaults  a system-wide defaults file, loaded over the shipped wopiserver.defaults.conf
//               (default /etc/wopi/wopiserver.defaults.conf, when present)

var configPath = "/etc/wopi/wopiserver.conf";
string? defaultsPath = null;
if (File.Exists("/etc/wopi/wopiserver.defaults.conf"))
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

    await Console.Error.WriteLineAsync($"Unknown argument {arg}. Usage: wopiserver [--config <path>] [--defaults <path>]");
    return 2;
}

WebApplication app;
try
{
    app = await WopiApp.Build(WopiConfig.Load(defaultsPath, configPath));
}
catch (Exception exception) when (exception is ConfigException or IOException or UnauthorizedAccessException or CryptographicException or FormatException or UnsupportedAppException)
{
    await Console.Error.WriteLineAsync($"Failed to initialize the service: {exception.Message}");
    return 1;
}

await app.RunAsync();
return 0;
