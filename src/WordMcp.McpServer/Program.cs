using System.Reflection;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;

namespace WordMcp.McpServer;

/// <summary>
/// WordMcp Model Context Protocol (MCP) server.
/// Exposes Microsoft Word automation as MCP tools over stdio.
/// </summary>
public static class Program
{
    /// <summary>
    /// Entry point.
    /// </summary>
    /// <param name="args">Command line arguments.</param>
    /// <returns>0 on success, 1 on a fatal error.</returns>
    public static async Task<int> Main(string[] args)
    {
        ArgumentNullException.ThrowIfNull(args);

        // office.dll (Microsoft.Office.Core) is a .NET Framework GAC assembly that .NET Core
        // cannot locate through standard probing.
        RegisterOfficeAssemblyResolver();

        if (args.Length > 0)
        {
            var arg = args[0].ToLowerInvariant();
            if (arg is "-h" or "--help" or "-?" or "/?" or "/h")
            {
                ShowHelp();
                return 0;
            }

            if (arg is "-v" or "--version")
            {
                Console.WriteLine($"Word MCP Server v{Version}");
                return 0;
            }
        }

        RegisterGlobalExceptionHandlers();

        var builder = Host.CreateApplicationBuilder(args);

        // Host.CreateApplicationBuilder enables reloadOnChange by default, which starts a
        // FileSystemWatcher. Word produces bursts of temp and lock files, which makes that
        // watcher spin. The server configuration never changes at runtime, so disable it.
        builder.Configuration.Sources.Clear();
        builder.Configuration
            .AddJsonFile("appsettings.json", optional: true, reloadOnChange: false)
            .AddEnvironmentVariables()
            .AddCommandLine(args);

        // With stdio transport, anything on stderr is surfaced as an error by MCP clients,
        // so only warnings and above are logged.
        builder.Logging.ClearProviders();
        builder.Logging.AddConsole(options => options.LogToStandardErrorThreshold = LogLevel.Warning);
        builder.Logging.SetMinimumLevel(LogLevel.Warning);

        builder.Services
            .AddMcpServer(options =>
            {
                options.ServerInfo = new()
                {
                    Name = "word-mcp",
                    Version = Version
                };

                options.ServerInstructions = """
                    WordMcp automates Microsoft Word via COM interop.

                    CRITICAL: the document must be CLOSED in the Word desktop app — COM requires
                    exclusive access.

                    SESSION LIFECYCLE:
                    1. file(action:'open', path:'C:\\...\\report.docx') -> returns session_id
                       (or file(action:'create', path:...) for a new document)
                    2. pass session_id to document / text / paragraph / table
                    3. file(action:'close', session_id:..., save:true) when completely done

                    Paragraph and table indexes are 1-based and shift after structural edits;
                    re-run paragraph(list) or table(list) before further edits.
                    """;
            })
            .WithToolsFromAssembly()
            .WithStdioServerTransport();

        var host = builder.Build();
        WordServices.Logger = host.Services.GetRequiredService<ILoggerFactory>().CreateLogger("WordMcp");

        try
        {
            await host.RunAsync().ConfigureAwait(false);
            return 0;
        }
        catch (OperationCanceledException)
        {
            // Graceful shutdown (Ctrl+C, SIGTERM).
            return 0;
        }
#pragma warning disable CA1031 // Top-level handler must not crash the process
        catch (Exception ex)
        {
            Console.Error.WriteLine($"[WordMcp] Fatal error: {ex.Message}");
            return 1;
        }
#pragma warning restore CA1031
        finally
        {
            // Without this, a client disconnect would silently discard unsaved work and
            // leave WINWORD.EXE running.
            WordServices.Shutdown();
        }
    }

    private static string Version
        => typeof(Program).Assembly.GetName().Version?.ToString() ?? "0.1.0";

    private static void RegisterGlobalExceptionHandlers()
    {
        AppDomain.CurrentDomain.UnhandledException += (_, e) =>
        {
            if (e.ExceptionObject is Exception ex)
            {
                Console.Error.WriteLine($"[WordMcp] Unhandled exception: {ex.Message}");
            }
        };

        TaskScheduler.UnobservedTaskException += (_, e) =>
            Console.Error.WriteLine($"[WordMcp] Unobserved task exception: {e.Exception.Message}");
    }

    private static void RegisterOfficeAssemblyResolver()
    {
        AppDomain.CurrentDomain.AssemblyResolve += (_, args) =>
        {
            var name = new AssemblyName(args.Name);
            return string.Equals(name.Name, "office", StringComparison.OrdinalIgnoreCase)
                ? ResolveOfficeDll()
                : null;
        };
    }

    /// <summary>
    /// Locates office.dll in the build output, the .NET Framework GAC, or an Office installation.
    /// </summary>
    private static Assembly? ResolveOfficeDll()
    {
        var localPath = Path.Combine(AppContext.BaseDirectory, "office.dll");
        if (File.Exists(localPath))
            return Assembly.LoadFrom(localPath);

        string[] gacPaths =
        [
            @"C:\Windows\assembly\GAC_MSIL\office\16.0.0.0__71e9bce111e9429c\OFFICE.DLL",
            @"C:\Windows\assembly\GAC_MSIL\office\15.0.0.0__71e9bce111e9429c\OFFICE.DLL",
        ];
        foreach (var gacPath in gacPaths)
        {
            if (File.Exists(gacPath))
                return Assembly.LoadFrom(gacPath);
        }

        string[] officeDirs =
        [
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), @"Microsoft Office\root\Office16"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), @"Microsoft Office\root\Office16"),
        ];
        foreach (var dir in officeDirs)
        {
            var officePath = Path.Combine(dir, "OFFICE.dll");
            if (File.Exists(officePath))
                return Assembly.LoadFrom(officePath);
        }

        return null;
    }

    private static void ShowHelp()
        => Console.WriteLine($"""
            Word MCP Server v{Version}

            An MCP (Model Context Protocol) server for Microsoft Word automation.

            Usage:
              WordMcp.McpServer.exe [options]

            Options:
              -h, --help      Show this help message
              -v, --version   Show version information

            Without options, starts the MCP server in stdio mode.

            Tools:
              file        open / create / save / close / list / test documents
              document    statistics, properties, export to PDF, save as
              text        read, append, find, replace, format
              paragraph   list, add, insert, delete, set style or alignment
              table       list, create, read, set cell, add or delete rows, set style

            Requirements:
              - Windows
              - Microsoft Word 2016 or later (desktop version)
            """);
}
