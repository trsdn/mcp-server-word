using System.Diagnostics;
using System.Text;
using Xunit;

namespace WordMcp.Service.Tests;

/// <summary>
/// Host and client over a real named pipe. None of these need Word: every command used here is
/// answered from the service's own state, which keeps the transport under test rather than COM.
/// </summary>
public sealed class ServiceHostTests : IDisposable
{
    private readonly CancellationTokenSource _cts = new(TimeSpan.FromMinutes(2));
    private readonly List<Task> _running = [];
    private readonly List<IDisposable> _disposables = [];

    public void Dispose()
    {
        _cts.Cancel();

        try
        {
            Task.WaitAll([.. _running], TimeSpan.FromSeconds(30));
        }
        catch (AggregateException)
        {
            // Cancellation during teardown.
        }

        foreach (var disposable in _disposables)
        {
            disposable.Dispose();
        }

        _cts.Dispose();
    }

    private static string NewPipeName() => ServiceSecurity.GetPrivatePipeName($"test-{Guid.NewGuid():N}");

    private ServiceHost StartHost(out WordMcpService service, TimeSpan? idleTimeout = null)
    {
        var pipeName = NewPipeName();
        service = new WordMcpService(idleTimeout: idleTimeout ?? TimeSpan.FromMinutes(30));
        _disposables.Add(service);

        var host = ServiceHost.TryCreate(service, pipeName);
        Assert.NotNull(host);
        _disposables.Add(host);
        _running.Add(host.RunAsync(_cts.Token));

        return host;
    }

    private ServiceClient ClientFor(ServiceHost host)
    {
        var client = new ServiceClient(host.PipeName, new ServiceClientOptions
        {
            AutoStart = false,
            ConnectTimeout = TimeSpan.FromSeconds(2),
            RequestTimeout = TimeSpan.FromSeconds(30)
        });

        _disposables.Add(client);
        return client;
    }

    private static async Task WaitUntilReadyAsync(ServiceClient client)
    {
        var deadline = DateTimeOffset.UtcNow.AddSeconds(20);

        while (DateTimeOffset.UtcNow < deadline)
        {
            if (await client.PingAsync())
            {
                return;
            }

            await Task.Delay(50);
        }

        Assert.Fail($"No service answered on '{client.PipeName}' within 20 seconds.");
    }

    [Fact]
    public async Task AHostedServiceAnswersAPing()
    {
        var host = StartHost(out _);
        var client = ClientFor(host);

        await WaitUntilReadyAsync(client);

        Assert.True(await client.PingAsync());
    }

    [Fact]
    public async Task ARequestIsAnsweredWithTheServicesResult()
    {
        var host = StartHost(out _);
        var client = ClientFor(host);
        await WaitUntilReadyAsync(client);

        var response = await client.SendAsync(new ServiceRequest { Command = "session.list" });

        Assert.True(response.Success, response.ErrorMessage);
        Assert.Contains("\"count\":0", response.Result!, StringComparison.Ordinal);
    }

    [Fact]
    public async Task AFailedCommandComesBackAsAnUnsuccessfulResponse()
    {
        var host = StartHost(out _);
        var client = ClientFor(host);
        await WaitUntilReadyAsync(client);

        var response = await client.SendAsync(new ServiceRequest { Command = "session.nonsense" });

        Assert.False(response.Success);
        Assert.Contains("session.nonsense", response.ErrorMessage!, StringComparison.Ordinal);
    }

    [Fact]
    public async Task ManyRequestsShareOneConnection()
    {
        var host = StartHost(out _);
        var client = ClientFor(host);
        await WaitUntilReadyAsync(client);

        using var pipe = ServiceSecurity.CreateClient(host.PipeName);
        await pipe.ConnectAsync(5000, _cts.Token);

        var writer = new StreamWriter(pipe, new UTF8Encoding(false), 4096, leaveOpen: true) { AutoFlush = true };
        using var reader = new StreamReader(pipe, new UTF8Encoding(false), false, 4096, leaveOpen: true);

        for (var i = 0; i < 5; i++)
        {
            await writer.WriteLineAsync(
                ServiceProtocol.Serialize(new ServiceRequest { Command = "service.ping" }).AsMemory(),
                _cts.Token);

            var line = await reader.ReadLineAsync(_cts.Token);
            Assert.NotNull(line);
            Assert.True(ServiceProtocol.Deserialize<ServiceResponse>(line)!.Success);
        }
    }

    [Fact]
    public async Task AMalformedLineIsRejectedWithoutDroppingTheConnection()
    {
        var host = StartHost(out _);
        var client = ClientFor(host);
        await WaitUntilReadyAsync(client);

        using var pipe = ServiceSecurity.CreateClient(host.PipeName);
        await pipe.ConnectAsync(5000, _cts.Token);

        var writer = new StreamWriter(pipe, new UTF8Encoding(false), 4096, leaveOpen: true) { AutoFlush = true };
        using var reader = new StreamReader(pipe, new UTF8Encoding(false), false, 4096, leaveOpen: true);

        await writer.WriteLineAsync("{ this is not json".AsMemory(), _cts.Token);
        var rejected = ServiceProtocol.Deserialize<ServiceResponse>((await reader.ReadLineAsync(_cts.Token))!)!;

        Assert.False(rejected.Success);
        Assert.Contains("Malformed", rejected.ErrorMessage!, StringComparison.Ordinal);

        await writer.WriteLineAsync(
            ServiceProtocol.Serialize(new ServiceRequest { Command = "service.ping" }).AsMemory(), _cts.Token);
        var recovered = ServiceProtocol.Deserialize<ServiceResponse>((await reader.ReadLineAsync(_cts.Token))!)!;

        Assert.True(recovered.Success);
    }

    [Fact]
    public async Task TwoClientsAreServedAtTheSameTime()
    {
        var host = StartHost(out _);
        var first = ClientFor(host);
        var second = ClientFor(host);
        await WaitUntilReadyAsync(first);

        var responses = await Task.WhenAll(
            first.SendAsync(new ServiceRequest { Command = "service.ping" }),
            second.SendAsync(new ServiceRequest { Command = "session.list" }));

        Assert.All(responses, r => Assert.True(r.Success, r.ErrorMessage));
    }

    [Fact]
    public void ASecondHostCannotClaimTheSamePipe()
    {
        var host = StartHost(out var service);

        using var duplicate = ServiceHost.TryCreate(service, host.PipeName);

        Assert.Null(duplicate);
    }

    [Fact]
    public void ThePipeNameIsFreedWhenTheHostIsDisposed()
    {
        var pipeName = NewPipeName();
        using var service = new WordMcpService();

        var first = ServiceHost.TryCreate(service, pipeName);
        Assert.NotNull(first);
        first.Dispose();

        using var second = ServiceHost.TryCreate(service, pipeName);
        Assert.NotNull(second);
    }

    [Fact]
    public async Task AShutdownRequestStopsTheHost()
    {
        var pipeName = NewPipeName();
        using var service = new WordMcpService();
        using var host = ServiceHost.TryCreate(service, pipeName)!;
        var running = host.RunAsync(_cts.Token);

        using var client = new ServiceClient(pipeName, new ServiceClientOptions { AutoStart = false });
        await WaitUntilReadyAsync(client);

        Assert.True((await client.SendAsync(new ServiceRequest { Command = "service.shutdown" })).Success);

        await running.WaitAsync(TimeSpan.FromSeconds(30));
        Assert.True(running.IsCompletedSuccessfully);
    }

    [Fact]
    public async Task AnIdleHostStopsOnItsOwn()
    {
        var pipeName = NewPipeName();
        using var service = new WordMcpService(idleTimeout: TimeSpan.Zero);
        using var host = ServiceHost.TryCreate(service, pipeName)!;

        var running = host.RunAsync(_cts.Token);

        await running.WaitAsync(TimeSpan.FromSeconds(30));
        Assert.True(running.IsCompletedSuccessfully);
    }

    [Fact]
    public async Task CancellingTheTokenStopsTheHost()
    {
        var pipeName = NewPipeName();
        using var service = new WordMcpService();
        using var host = ServiceHost.TryCreate(service, pipeName)!;
        using var stopping = new CancellationTokenSource();

        var running = host.RunAsync(stopping.Token);
        using var client = new ServiceClient(pipeName, new ServiceClientOptions { AutoStart = false });
        await WaitUntilReadyAsync(client);

        await stopping.CancelAsync();

        await running.WaitAsync(TimeSpan.FromSeconds(30));
    }

    [Fact]
    public void TryCreate_RejectsMissingArguments()
    {
        using var service = new WordMcpService();

        Assert.Throws<ArgumentNullException>(() => ServiceHost.TryCreate(null!, "pipe"));
        Assert.Throws<ArgumentException>(() => ServiceHost.TryCreate(service, "  "));
    }

    [Fact]
    public async Task AClientReportsAMissingServiceWhenAutoStartIsOff()
    {
        using var client = new ServiceClient(NewPipeName(), new ServiceClientOptions
        {
            AutoStart = false,
            ConnectTimeout = TimeSpan.FromMilliseconds(500)
        });

        Assert.False(await client.PingAsync());

        var response = await client.SendAsync(new ServiceRequest { Command = "service.ping" });

        Assert.False(response.Success);
        Assert.Contains("No Word session service", response.ErrorMessage!, StringComparison.Ordinal);
    }

    [Fact]
    public async Task AClientReportsAFailedAutoStartWhenTheExecutableIsMissing()
    {
        using var client = new ServiceClient(NewPipeName(), new ServiceClientOptions
        {
            AutoStart = true,
            ConnectTimeout = TimeSpan.FromMilliseconds(500),
            ExecutablePath = Path.Combine(Path.GetTempPath(), $"missing-{Guid.NewGuid():N}.exe")
        });

        var response = await client.SendAsync(new ServiceRequest { Command = "service.ping" });

        Assert.False(response.Success);
        Assert.Contains("Could not start", response.ErrorMessage!, StringComparison.Ordinal);
    }

    [Fact]
    public void TheDaemonExecutableShipsNextToItsClients()
    {
        var executable = ServiceClient.FindExecutable();

        Assert.NotNull(executable);
        Assert.True(File.Exists(executable));
    }

    [Fact]
    public async Task AClientStartsTheDaemonItCannotReach()
    {
        var pipeName = NewPipeName();
        using var client = new ServiceClient(pipeName, new ServiceClientOptions
        {
            AutoStart = true,
            ConnectTimeout = TimeSpan.FromSeconds(1),
            StartupTimeout = TimeSpan.FromSeconds(45),
            IdleTimeout = TimeSpan.FromMinutes(1)
        });

        try
        {
            var response = await client.SendAsync(new ServiceRequest { Command = "service.ping" });

            Assert.True(response.Success, response.ErrorMessage);

            // The answer must come from another process, otherwise nothing was really started.
            var payload = System.Text.Json.JsonDocument.Parse(response.Result!).RootElement;
            Assert.NotEqual(Environment.ProcessId, payload.GetProperty("processId").GetInt32());
            Assert.False(Process.GetProcessById(payload.GetProperty("processId").GetInt32()).HasExited);
        }
        finally
        {
            await client.SendAsync(new ServiceRequest { Command = "service.shutdown" });
        }
    }

    [Fact]
    public async Task ASecondCallReusesTheDaemonTheFirstOneStarted()
    {
        var pipeName = NewPipeName();
        using var client = new ServiceClient(pipeName, new ServiceClientOptions
        {
            AutoStart = true,
            ConnectTimeout = TimeSpan.FromSeconds(1),
            StartupTimeout = TimeSpan.FromSeconds(45),
            IdleTimeout = TimeSpan.FromMinutes(1)
        });

        try
        {
            var first = await client.SendAsync(new ServiceRequest { Command = "service.ping" });
            var second = await client.SendAsync(new ServiceRequest { Command = "service.ping" });

            Assert.True(first.Success, first.ErrorMessage);
            Assert.True(second.Success, second.ErrorMessage);
            Assert.Equal(first.Result, second.Result);
        }
        finally
        {
            await client.SendAsync(new ServiceRequest { Command = "service.shutdown" });
        }
    }

    [Fact]
    public void AClientRejectsABlankPipeName()
        => Assert.Throws<ArgumentException>(() => new ServiceClient("  "));

    [Fact]
    public async Task ADisposedClientRefusesFurtherCalls()
    {
        var client = new ServiceClient(NewPipeName());
        client.Dispose();

        await Assert.ThrowsAsync<ObjectDisposedException>(
            () => client.SendAsync(new ServiceRequest { Command = "service.ping" }));
    }
}
