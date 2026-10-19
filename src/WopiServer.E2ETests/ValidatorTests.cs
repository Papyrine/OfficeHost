using System.Xml.Linq;

/// <summary>
/// Microsoft's WOPI validator (https://github.com/microsoft/wopi-validator-core) against the WOPI endpoints: each test
/// group on a fresh storage holding a .wopitest file. WOPI_VALIDATOR is the path of the validator's
/// Microsoft.Office.WopiValidator.dll, which src/wopi-validator.ps1 builds.
/// The ProofKeys group is not run: the validator's command line has no proof keys, so it sends no X-WOPI-Proof, and the
/// tests expecting a rejection cannot pass. ProofKeyTests covers the proof keys instead.
/// </summary>
public partial class ValidatorTests
{
    static readonly string[] excludedGroups = ["ProofKeys"];

    [System.Text.RegularExpressions.GeneratedRegex(@"(?m)^\s+Pass: ")]
    private static partial System.Text.RegularExpressions.Regex Passed();

    [System.Text.RegularExpressions.GeneratedRegex(@"(?m)^\s+Fail: (\S+)")]
    private static partial System.Text.RegularExpressions.Regex Failed();

    [Test]
    [Timeout(600_000)]
    public async Task WopiValidator(Cancel cancel)
    {
        var validator = Environment.GetEnvironmentVariable("WOPI_VALIDATOR");
        if (string.IsNullOrEmpty(validator))
        {
            Skip.Test("WOPI_VALIDATOR is not set: run src/wopi-validator.ps1");
        }

        var validatorFolder = Path.GetDirectoryName(validator)!;
        var groups = XDocument.Load(Path.Combine(validatorFolder, "TestCases.xml"))
            .Descendants("TestGroup")
            .Select(_ => (string) _.Attribute("Name")!)
            .Except(excludedGroups)
            .ToList();
        var passed = 0;
        var failed = new List<string>();
        foreach (var group in groups)
        {
            await using var server = await RunningServer.Start(host: "localhost");
            await File.WriteAllBytesAsync(Path.Combine(server.StoragePath, "test.wopitest"), [], cancel);
            var (fileId, token) = await server.Open("test.wopitest");
            var wopiSrc = WopiSrc(server, fileId);
            var (code, output) = await Run(validatorFolder, ["exec", validator, "--wopisrc", wopiSrc, "--token", token, "--token_ttl", "0", "--ignore-skipped", "--testgroup", group], cancel);
            Console.WriteLine(output.TrimEnd());
            passed += Passed().Count(output);
            var groupFailures = Failed().Matches(output).Select(_ => $"{group}: {_.Groups[1].Value}").ToList();
            if (code != 0 &&
                groupFailures.Count == 0)
            {
                // the validator failed without reporting a test, for instance on an exception
                groupFailures.Add($"{group} (exit code {code})");
            }

            if (groupFailures.Count > 0)
            {
                Console.WriteLine(string.Join('\n', server.Requests));
                Console.WriteLine(server.Log.ToString());
            }

            failed.AddRange(groupFailures);
        }

        Console.WriteLine($"WOPI validator: {passed} passed, {failed.Count} failed (excluded groups: {string.Join(", ", excludedGroups)})");
        await Assert.That(failed).IsEmpty();
        await Assert.That(passed).IsGreaterThan(0);
    }

    static string WopiSrc(RunningServer server, string fileId)
    {
        var baseUrl = server.Service<Microsoft.Extensions.Options.IOptions<WopiOptions>>().Value.BaseUrl;
        return $"{baseUrl}/wopi/files/{Uri.EscapeDataString(fileId)}";
    }

    /// <summary>
    /// Runs dotnet in the folder of the validator, which reads its test definitions (TestCases.xml) from the current
    /// directory.
    /// </summary>
    static async Task<(int Code, string Output)> Run(string folder, IEnumerable<string> args, Cancel cancel)
    {
        var start = new ProcessStartInfo("dotnet")
        {
            WorkingDirectory = folder,
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
