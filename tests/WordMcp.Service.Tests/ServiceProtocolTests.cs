using System.Text.Json;
using Xunit;

namespace WordMcp.Service.Tests;

/// <summary>
/// Protocol tests. The wire format is line delimited, so the property these assert is that a
/// serialized message never contains a raw newline no matter what the payload holds.
/// </summary>
public class ServiceProtocolTests
{
    [Fact]
    public void Serialize_ProducesASingleLine()
    {
        var request = new ServiceRequest
        {
            Command = "session.open",
            Args = ServiceProtocol.Serialize(new { filePath = "C:\\docs\\a\r\nb.docx" })
        };

        var line = ServiceProtocol.Serialize(request);

        Assert.DoesNotContain("\n", line, StringComparison.Ordinal);
        Assert.DoesNotContain("\r", line, StringComparison.Ordinal);
    }

    [Fact]
    public void Serialize_RoundTripsThroughDeserialize()
    {
        var request = new ServiceRequest
        {
            Command = "session.close",
            SessionId = "word-abc",
            Args = "{\"save\":true}",
            Source = "mcp"
        };

        var restored = ServiceProtocol.Deserialize<ServiceRequest>(ServiceProtocol.Serialize(request));

        Assert.NotNull(restored);
        Assert.Equal("session.close", restored.Command);
        Assert.Equal("word-abc", restored.SessionId);
        Assert.Equal("{\"save\":true}", restored.Args);
        Assert.Equal("mcp", restored.Source);
    }

    [Fact]
    public void Serialize_UsesCamelCaseAndOmitsNulls()
    {
        var line = ServiceProtocol.Serialize(new ServiceRequest { Command = "service.ping" });

        Assert.Contains("\"command\":\"service.ping\"", line, StringComparison.Ordinal);
        Assert.DoesNotContain("sessionId", line, StringComparison.Ordinal);
        Assert.DoesNotContain("null", line, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void Deserialize_TreatsABlankLineAsNoMessage(string line)
        => Assert.Null(ServiceProtocol.Deserialize<ServiceRequest>(line));

    [Fact]
    public void Deserialize_RejectsMalformedJson()
        => Assert.Throws<JsonException>(() => ServiceProtocol.Deserialize<ServiceRequest>("{not json"));

    [Fact]
    public void Ok_SerializesThePayloadIntoResult()
    {
        var response = ServiceResponse.Ok(new { sessionId = "word-1", count = 2 });

        Assert.True(response.Success);
        Assert.Null(response.ErrorMessage);
        Assert.Equal("{\"sessionId\":\"word-1\",\"count\":2}", response.Result);
    }

    [Fact]
    public void Fail_CarriesTheMessageAndType()
    {
        var response = ServiceResponse.Fail("nope", nameof(InvalidOperationException));

        Assert.False(response.Success);
        Assert.Equal("nope", response.ErrorMessage);
        Assert.Equal(nameof(InvalidOperationException), response.ErrorType);
        Assert.Null(response.Result);
    }

    [Fact]
    public void Status_RoundTrips()
    {
        var status = new ServiceStatus
        {
            ProcessId = 4711,
            SessionCount = 3,
            StartedAt = DateTimeOffset.UnixEpoch,
            LastActivityAt = DateTimeOffset.UnixEpoch.AddMinutes(5),
            IdleTimeout = TimeSpan.FromMinutes(30),
            Version = "1.2.3"
        };

        var restored = ServiceProtocol.Deserialize<ServiceStatus>(ServiceProtocol.Serialize(status));

        Assert.NotNull(restored);
        Assert.Equal(4711, restored.ProcessId);
        Assert.Equal(3, restored.SessionCount);
        Assert.Equal(TimeSpan.FromMinutes(30), restored.IdleTimeout);
        Assert.Equal("1.2.3", restored.Version);
    }
}
