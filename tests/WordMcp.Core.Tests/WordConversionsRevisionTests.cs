using WordMcp.ComInterop;
using WordMcp.Core.Utilities;
using Xunit;

namespace WordMcp.Core.Tests;

/// <summary>
/// Conversion of Word's revision type constants to the friendly names on the wire.
/// </summary>
public class WordConversionsRevisionTests
{
    [Theory]
    [InlineData(ComInteropConstants.WdNoRevision, "none")]
    [InlineData(ComInteropConstants.WdRevisionInsert, "insert")]
    [InlineData(ComInteropConstants.WdRevisionDelete, "delete")]
    [InlineData(ComInteropConstants.WdRevisionReplace, "replace")]
    [InlineData(ComInteropConstants.WdRevisionMovedFrom, "moved-from")]
    [InlineData(ComInteropConstants.WdRevisionMovedTo, "moved-to")]
    public void FromWdRevisionType_MapsDistinctKinds(int wdRevisionType, string expected)
        => Assert.Equal(expected, WordConversions.FromWdRevisionType(wdRevisionType));

    [Theory]
    [InlineData(ComInteropConstants.WdRevisionProperty)]
    [InlineData(ComInteropConstants.WdRevisionStyle)]
    [InlineData(ComInteropConstants.WdRevisionParagraphProperty)]
    [InlineData(ComInteropConstants.WdRevisionTableProperty)]
    [InlineData(ComInteropConstants.WdRevisionSectionProperty)]
    [InlineData(ComInteropConstants.WdRevisionStyleDefinition)]
    public void FromWdRevisionType_CollapsesFormattingKinds(int wdRevisionType)
        => Assert.Equal("format", WordConversions.FromWdRevisionType(wdRevisionType));

    [Fact]
    public void FromWdRevisionType_FallsBackForUnknownConstants()
        => Assert.Equal("other", WordConversions.FromWdRevisionType(99));
}
