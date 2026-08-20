using WordMcp.Core.Commands.List;
using Xunit;

namespace WordMcp.Core.Tests;

/// <summary>
/// Argument validation of the list commands. Everything asserted here happens before the batch is
/// touched, so these tests need no Word installation and run in CI.
/// </summary>
public class ListValidationTests
{
    private static readonly ListCommands Lists = new();
    private static readonly ThrowingBatch Batch = new();

    [Fact]
    public void Apply_RejectsStartIndexBelowOne()
        => Assert.Throws<ArgumentOutOfRangeException>(() => Lists.Apply(Batch, 0));

    [Fact]
    public void Apply_RejectsUnknownListType()
        => Assert.Throws<ArgumentException>(() => Lists.Apply(Batch, 1, listType: "squiggle"));

    [Theory]
    [InlineData("bullet")]
    [InlineData("number")]
    [InlineData("outline-number")]
    public void Apply_AcceptsKnownListTypes(string listType)
        => Assert.Throws<NotSupportedException>(() => Lists.Apply(Batch, 1, listType: listType));

    [Theory]
    [InlineData(0)]
    [InlineData(10)]
    public void Apply_RejectsLevelOutsideOneToNine(int level)
        => Assert.Throws<ArgumentOutOfRangeException>(() => Lists.Apply(Batch, 1, level: level));

    [Fact]
    public void SetLevel_RejectsStartIndexBelowOne()
        => Assert.Throws<ArgumentOutOfRangeException>(() => Lists.SetLevel(Batch, 0, 1));

    [Theory]
    [InlineData(0)]
    [InlineData(10)]
    public void SetLevel_RejectsLevelOutsideOneToNine(int level)
        => Assert.Throws<ArgumentOutOfRangeException>(() => Lists.SetLevel(Batch, 1, level));

    [Fact]
    public void SetLevel_ReachesBatchForValidArguments()
        => Assert.Throws<NotSupportedException>(() => Lists.SetLevel(Batch, 1, 9));

    [Fact]
    public void Restart_RejectsStartIndexBelowOne()
        => Assert.Throws<ArgumentOutOfRangeException>(() => Lists.Restart(Batch, 0));

    [Fact]
    public void Remove_RejectsStartIndexBelowOne()
        => Assert.Throws<ArgumentOutOfRangeException>(() => Lists.Remove(Batch, -1));

    [Fact]
    public void Remove_ReachesBatchForValidArguments()
        => Assert.Throws<NotSupportedException>(() => Lists.Remove(Batch, 1, 4));

    [Fact]
    public void AllCommands_RejectNullBatch()
    {
        Assert.Throws<ArgumentNullException>(() => Lists.Get(null!));
        Assert.Throws<ArgumentNullException>(() => Lists.Apply(null!, 1));
        Assert.Throws<ArgumentNullException>(() => Lists.SetLevel(null!, 1, 1));
        Assert.Throws<ArgumentNullException>(() => Lists.Restart(null!, 1));
        Assert.Throws<ArgumentNullException>(() => Lists.Remove(null!, 1));
    }
}
