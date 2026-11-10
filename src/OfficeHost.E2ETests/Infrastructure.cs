/// <summary>
/// A WOPI client running in Docker.
/// </summary>
sealed class EditorContainer(string name, int port) : IAsyncDisposable
{
    /// <summary>
    /// The port of the editor on this machine.
    /// </summary>
    public int Port { get; } = port;

    public string Url => $"http://localhost:{Port}";

    public string DiscoveryUrl => $"{Url}/hosting/discovery";

    /// <summary>
    /// Starts a container and waits until its WOPI discovery answers.
    /// </summary>
    public static async Task<EditorContainer> Start(string image, int port, int containerPort, IReadOnlyDictionary<string, string> environment, TimeSpan startTimeout, Cancel cancel)
    {
        var name = $"wopiserver-e2e-{Guid.NewGuid().ToString("N")[..8]}";
        List<string> args =
        [
            "run", "--detach", "--rm", "--name", name, "--publish", $"{port}:{containerPort}",
            // on Linux, host.docker.internal is not wired by default: map it to the host
            "--add-host", "host.docker.internal:host-gateway"
        ];
        foreach (var (key, value) in environment)
        {
            args.AddRange(["--env", $"{key}={value}"]);
        }

        args.Add(image);
        var (code, output) = await Docker(args, cancel);
        if (code != 0)
        {
            throw new($"docker run failed: {output}");
        }

        var container = new EditorContainer(name, port);
        using var http = new HttpClient();
        var deadline = DateTime.UtcNow + startTimeout;
        while (true)
        {
            try
            {
                using var response = await http.GetAsync($"{container.Url}/hosting/discovery", cancel);
                if (response.IsSuccessStatusCode)
                {
                    return container;
                }
            }
            catch (HttpRequestException)
            {
            }

            if (DateTime.UtcNow > deadline)
            {
                var logs = await container.Logs(cancel);
                await container.DisposeAsync();
                throw new($"{image} did not start in {startTimeout}: {logs}");
            }

            await Task.Delay(2000, cancel);
        }
    }

    public async Task<string> Logs(Cancel cancel) =>
        (await Docker(["logs", "--tail", "200", name], cancel)).Output;

    public async ValueTask DisposeAsync() =>
        await Docker(["rm", "--force", name], Cancel.None);

    static async Task<(int Code, string Output)> Docker(IEnumerable<string> args, Cancel cancel)
    {
        var start = new ProcessStartInfo("docker")
        {
            RedirectStandardOutput = true,
            RedirectStandardError = true
        };
        foreach (var arg in args)
        {
            start.ArgumentList.Add(arg);
        }

        using var process = Process.Start(start)!;
        var output = process.StandardOutput.ReadToEndAsync(cancel);
        var error = process.StandardError.ReadToEndAsync(cancel);
        await process.WaitForExitAsync(cancel);
        return (process.ExitCode, $"{await output}{await error}");
    }
}