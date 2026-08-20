using Xunit;

namespace WordMcp.Service.Tests;

/// <summary>
/// Command line parsing. Unknown input selects help rather than failing, so these pin that down
/// alongside the ordinary modes.
/// </summary>
public class CommandLineTests
{    [Fact]
    public void NoArgumentsShowsHelp()
        => Assert.Equal(RunMode.Help, CommandLine.Parse([]).Mode);

    [Theory]
    [InlineData("--daemon", RunMode.Daemon)]
    [InlineData("--status", RunMode.Status)]
    [InlineData("--stop", RunMode.Stop)]
    [InlineData("--version", RunMode.Version)]
    [InlineData("-v", RunMode.Version)]
    [InlineData("--help", RunMode.Help)]
    [InlineData("-h", RunMode.Help)]
    public void AModeIsRecognised(string argument, RunMode expected)
        => Assert.Equal(expected, CommandLine.Parse([argument]).Mode);

    [Fact]
    public void ModesAreCaseInsensitive()
        => Assert.Equal(RunMode.Daemon, CommandLine.Parse(["--DAEMON"]).Mode);

    [Fact]
    public void AnUnknownArgumentFallsBackToHelp()
        => Assert.Equal(RunMode.Help, CommandLine.Parse(["--frobnicate"]).Mode);

    [Fact]
    public void ThePipeDefaultsToThePerUserServicePipe()
        => Assert.Equal(ServiceSecurity.GetServicePipeName(), CommandLine.Parse(["--daemon"]).PipeName);

    [Fact]
    public void ThePipeCanBeOverridden()
        => Assert.Equal("custom-pipe", CommandLine.Parse(["--daemon", "--pipe", "custom-pipe"]).PipeName);

    [Fact]
    public void ADanglingPipeArgumentKeepsTheDefault()
        => Assert.Equal(ServiceSecurity.GetServicePipeName(), CommandLine.Parse(["--daemon", "--pipe"]).PipeName);

    [Fact]
    public void TheIdleTimeoutIsReadInMinutes()
        => Assert.Equal(TimeSpan.FromMinutes(7), CommandLine.Parse(["--daemon", "--idle-minutes", "7"]).IdleTimeout);

    [Fact]
    public void AnIdleTimeoutOfZeroIsAllowed()
        => Assert.Equal(TimeSpan.Zero, CommandLine.Parse(["--daemon", "--idle-minutes", "0"]).IdleTimeout);

    [Theory]
    [InlineData("-5")]
    [InlineData("soon")]
    public void AnUnusableIdleTimeoutLeavesTheDefaultInPlace(string value)
        => Assert.Null(CommandLine.Parse(["--daemon", "--idle-minutes", value]).IdleTimeout);

    [Fact]
    public void VerboseIsOffByDefault()
        => Assert.False(CommandLine.Parse(["--daemon"]).Verbose);

    [Fact]
    public void VerboseCanBeTurnedOn()
        => Assert.True(CommandLine.Parse(["--daemon", "--verbose"]).Verbose);

    [Fact]
    public void TheLastModeWins()
        => Assert.Equal(RunMode.Stop, CommandLine.Parse(["--daemon", "--stop"]).Mode);
}
