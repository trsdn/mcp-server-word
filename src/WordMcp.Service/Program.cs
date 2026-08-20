using System.Globalization;
using Microsoft.Extensions.Logging;

namespace WordMcp.Service;

/// <summary>
/// Command line entry point of the Word session service.
/// </summary>
public static class Program
{
    /// <summary>
    /// Runs the requested mode.
    /// </summary>
    /// <param name="args">Command line arguments.</param>
    /// <returns>0 on success, 1 on failure.</returns>
    public static async Task<int> Main(string[] args)
    {
        ArgumentNullException.ThrowIfNull(args);

        var options = CommandLine.Parse(args);

        return options.Mode switch
        {
            RunMode.Help => Show(Usage()),
            RunMode.Version => Show($"Word session service v{VersionText}"),
            RunMode.Daemon => await RunDaemonAsync(options).ConfigureAwait(false),
            RunMode.Status => await ShowStatusAsync(options).ConfigureAwait(false),
            RunMode.Stop => await StopAsync(options).ConfigureAwait(false),
            _ => Show(Usage())
        };
    }

    private static string VersionText
        => typeof(Program).Assembly.GetName().Version?.ToString() ?? "0.1.0";

    private static int Show(string text)
    {
        Console.WriteLine(text);
        return 0;
    }

    private static async Task<int> RunDaemonAsync(CommandLineOptions options)
    {
        using var loggerFactory = LoggerFactory.Create(builder => builder
            .AddSimpleConsole(console => console.SingleLine = true)
            .SetMinimumLevel(options.Verbose ? LogLevel.Information : LogLevel.Warning));

        var logger = loggerFactory.CreateLogger("WordMcp.Service");

        using var service = new WordMcpService(logger, options.IdleTimeout);
        using var host = ServiceHost.TryCreate(service, options.PipeName, logger);

        if (host is null)
        {
            // Not an error: the caller wanted a service on that pipe and there is one.
            Console.Error.WriteLine($"A Word session service is already listening on '{options.PipeName}'.");
            return 0;
        }

        using var stopping = new CancellationTokenSource();
        Console.CancelKeyPress += (_, e) =>
        {
            e.Cancel = true;
            stopping.Cancel();
        };

        try
        {
            await host.RunAsync(stopping.Token).ConfigureAwait(false);
            return 0;
        }
        catch (OperationCanceledException)
        {
            return 0;
        }
#pragma warning disable CA1031 // Top-level handler must not crash the process
        catch (Exception ex)
        {
            Console.Error.WriteLine($"[WordMcp.Service] Fatal error: {ex.Message}");
            return 1;
        }
#pragma warning restore CA1031
    }

    private static async Task<int> ShowStatusAsync(CommandLineOptions options)
    {
        using var client = new ServiceClient(options.PipeName, new ServiceClientOptions { AutoStart = false });

        var response = await client.SendAsync(new ServiceRequest { Command = "service.status", Source = "cli" })
            .ConfigureAwait(false);

        if (!response.Success)
        {
            Console.WriteLine($"Not running on '{options.PipeName}'.");
            return 1;
        }

        var status = ServiceProtocol.Deserialize<ServiceStatus>(response.Result!)!;

        Console.WriteLine(string.Create(CultureInfo.InvariantCulture, $"""
            Word session service v{status.Version}
              pipe          {options.PipeName}
              process       {status.ProcessId}
              sessions      {status.SessionCount}
              started       {status.StartedAt:u}
              last activity {status.LastActivityAt:u}
              idle timeout  {status.IdleTimeout}
            """));

        return 0;
    }

    private static async Task<int> StopAsync(CommandLineOptions options)
    {
        using var client = new ServiceClient(options.PipeName, new ServiceClientOptions { AutoStart = false });

        var response = await client.SendAsync(new ServiceRequest { Command = "service.shutdown", Source = "cli" })
            .ConfigureAwait(false);

        Console.WriteLine(response.Success
            ? "The Word session service is shutting down; open documents are saved first."
            : $"Not running on '{options.PipeName}'.");

        return response.Success ? 0 : 1;
    }

    private static string Usage() => $"""
        Word session service v{VersionText}

        Holds Microsoft Word sessions so several clients can share one open document.

        Usage:
          WordMcp.Service.exe --daemon [--pipe NAME] [--idle-minutes N] [--verbose]
          WordMcp.Service.exe --status [--pipe NAME]
          WordMcp.Service.exe --stop   [--pipe NAME]

        Options:
          --daemon          Listen for clients until idle or stopped
          --status          Print what a running service is doing
          --stop            Ask a running service to save and exit
          --pipe NAME       Pipe to use (default: the per-user service pipe)
          --idle-minutes N  Exit after N minutes without sessions (default: {(int)WordMcpService.DefaultIdleTimeout.TotalMinutes})
          --verbose         Log informational messages as well as warnings
          -h, --help        Show this help
          -v, --version     Show version information

        The pipe name embeds the current user's SID and is ACL'd to that user, so
        sessions are never shared across accounts.
        """;
}

/// <summary>
/// What the process was asked to do.
/// </summary>
public enum RunMode
{
    /// <summary>Print usage.</summary>
    Help,

    /// <summary>Print the version.</summary>
    Version,

    /// <summary>Listen for clients.</summary>
    Daemon,

    /// <summary>Report what a running service is doing.</summary>
    Status,

    /// <summary>Ask a running service to save and exit.</summary>
    Stop
}

/// <summary>
/// Parsed command line.
/// </summary>
public sealed class CommandLineOptions
{
    /// <summary>
    /// Gets the requested mode.
    /// </summary>
    public RunMode Mode { get; init; }

    /// <summary>
    /// Gets the pipe to listen on or talk to.
    /// </summary>
    public required string PipeName { get; init; }

    /// <summary>
    /// Gets the idle timeout, or <c>null</c> to use the service default.
    /// </summary>
    public TimeSpan? IdleTimeout { get; init; }

    /// <summary>
    /// Gets a value indicating whether informational messages are logged.
    /// </summary>
    public bool Verbose { get; init; }
}

/// <summary>
/// Turns arguments into <see cref="CommandLineOptions"/>.
/// </summary>
public static class CommandLine
{
    /// <summary>
    /// Parses arguments. Unknown arguments select the help mode rather than failing, because a
    /// service that refuses to explain itself is worse than one that prints usage.
    /// </summary>
    /// <param name="args">Raw arguments.</param>
    /// <returns>The parsed options.</returns>
    public static CommandLineOptions Parse(string[] args)
    {
        var mode = RunMode.Help;
        string? pipeName = null;
        TimeSpan? idleTimeout = null;
        var verbose = false;

        for (var i = 0; i < args.Length; i++)
        {
            switch (args[i].ToLowerInvariant())
            {
                case "--daemon":
                    mode = RunMode.Daemon;
                    break;
                case "--status":
                    mode = RunMode.Status;
                    break;
                case "--stop":
                    mode = RunMode.Stop;
                    break;
                case "-v":
                case "--version":
                    mode = RunMode.Version;
                    break;
                case "-h":
                case "--help":
                case "-?":
                case "/?":
                    mode = RunMode.Help;
                    break;
                case "--verbose":
                    verbose = true;
                    break;
                case "--pipe" when i + 1 < args.Length:
                    pipeName = args[++i];
                    break;
                case "--idle-minutes" when i + 1 < args.Length:
                    if (int.TryParse(args[++i], NumberStyles.Integer, CultureInfo.InvariantCulture, out var minutes) &&
                        minutes >= 0)
                    {
                        idleTimeout = TimeSpan.FromMinutes(minutes);
                    }

                    break;
                default:
                    break;
            }
        }

        return new CommandLineOptions
        {
            Mode = mode,
            PipeName = string.IsNullOrWhiteSpace(pipeName) ? ServiceSecurity.GetServicePipeName() : pipeName,
            IdleTimeout = idleTimeout,
            Verbose = verbose
        };
    }
}
