using System.Security.Principal;
using Xunit;

namespace WordMcp.Service.Tests;

/// <summary>
/// Pipe naming and access control. The security promise is that a pipe is both named per user
/// and ACL'd to that user, so these check the name and then actually open one end to end.
/// </summary>
public class ServiceSecurityTests
{
    [Fact]
    public void ServicePipeName_ContainsTheCurrentUserSid()
    {
        var sid = WindowsIdentity.GetCurrent().User!.Value;

        Assert.Contains(sid, ServiceSecurity.GetServicePipeName(), StringComparison.Ordinal);
    }

    [Fact]
    public void ServicePipeName_IsStableAcrossCalls()
        => Assert.Equal(ServiceSecurity.GetServicePipeName(), ServiceSecurity.GetServicePipeName());

    [Fact]
    public void PrivatePipeName_DiffersPerInstance()
        => Assert.NotEqual(
            ServiceSecurity.GetPrivatePipeName("a"),
            ServiceSecurity.GetPrivatePipeName("b"));

    [Fact]
    public void PrivatePipeName_DiffersFromTheSharedName()
        => Assert.NotEqual(
            ServiceSecurity.GetServicePipeName(),
            ServiceSecurity.GetPrivatePipeName("a"));

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void PrivatePipeName_RejectsABlankInstance(string instance)
        => Assert.Throws<ArgumentException>(() => ServiceSecurity.GetPrivatePipeName(instance));

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void CreateSecureServer_RejectsABlankName(string name)
        => Assert.Throws<ArgumentException>(() => ServiceSecurity.CreateSecureServer(name));

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void CreateClient_RejectsABlankName(string name)
        => Assert.Throws<ArgumentException>(() => ServiceSecurity.CreateClient(name));

    [Fact]
    public async Task AClientOfTheSameUserCanConnectToASecureServer()
    {
        var pipeName = ServiceSecurity.GetPrivatePipeName($"connect-{Guid.NewGuid():N}");
        using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(15));

        using var server = ServiceSecurity.CreateSecureServer(pipeName);
        var accepting = server.WaitForConnectionAsync(cts.Token);

        using var client = ServiceSecurity.CreateClient(pipeName);
        await client.ConnectAsync(5000, cts.Token);
        await accepting;

        Assert.True(server.IsConnected);
        Assert.True(client.IsConnected);
    }

    [Fact]
    public async Task ASecurePipeCarriesBytesInBothDirections()
    {
        var pipeName = ServiceSecurity.GetPrivatePipeName($"echo-{Guid.NewGuid():N}");
        using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(15));
        var token = cts.Token;

        using var server = ServiceSecurity.CreateSecureServer(pipeName);
        var accepting = server.WaitForConnectionAsync(token);

        using var client = ServiceSecurity.CreateClient(pipeName);
        await client.ConnectAsync(5000, token);
        await accepting;

        var writer = new StreamWriter(server, leaveOpen: true) { AutoFlush = true };
        await writer.WriteLineAsync("hello".AsMemory(), token);

        using var reader = new StreamReader(client, leaveOpen: true);
        Assert.Equal("hello", await reader.ReadLineAsync(token));
    }
}
