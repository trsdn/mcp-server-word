using WordMcp.ComInterop;
using WordMcp.Core.Utilities;
using Xunit;

namespace WordMcp.Core.Tests;

/// <summary>
/// Conversions between the friendly style-type names on the wire and Word's WdStyleType constants.
/// </summary>
public class WordConversionsStyleTests
{
    [Theory]
    [InlineData("paragraph", ComInteropConstants.WdStyleTypeParagraph)]
    [InlineData("character", ComInteropConstants.WdStyleTypeCharacter)]
    [InlineData("table", ComInteropConstants.WdStyleTypeTable)]
    [InlineData("list", ComInteropConstants.WdStyleTypeList)]
    public void ToWdStyleType_MapsKnownNames(string styleType, int expected)
        => Assert.Equal(expected, WordConversions.ToWdStyleType(styleType));

    [Theory]
    [InlineData("PARAGRAPH")]
    [InlineData("  Character  ")]
    public void ToWdStyleType_IgnoresCaseAndSurroundingSpace(string styleType)
        => Assert.True(WordConversions.ToWdStyleType(styleType) > 0);

    [Fact]
    public void ToWdStyleType_RejectsUnknownName()
        => Assert.Throws<ArgumentException>(() => WordConversions.ToWdStyleType("sideways"));

    [Fact]
    public void ToWdStyleType_RejectsEmptyName()
        => Assert.Throws<ArgumentException>(() => WordConversions.ToWdStyleType("  "));

    [Theory]
    [InlineData(ComInteropConstants.WdStyleTypeParagraph, "paragraph")]
    [InlineData(ComInteropConstants.WdStyleTypeCharacter, "character")]
    [InlineData(ComInteropConstants.WdStyleTypeTable, "table")]
    [InlineData(ComInteropConstants.WdStyleTypeList, "list")]
    public void FromWdStyleType_MapsKnownConstants(int wdStyleType, string expected)
        => Assert.Equal(expected, WordConversions.FromWdStyleType(wdStyleType));

    [Fact]
    public void FromWdStyleType_FallsBackForUnknownConstants()
        => Assert.Equal("other", WordConversions.FromWdStyleType(99));

    [Theory]
    [InlineData("paragraph")]
    [InlineData("character")]
    [InlineData("table")]
    [InlineData("list")]
    public void StyleTypeConversion_RoundTrips(string styleType)
        => Assert.Equal(styleType, WordConversions.FromWdStyleType(WordConversions.ToWdStyleType(styleType)));

    [Fact]
    public void KnownBuiltInStyles_CoversTheStylesWordStylesResolves()
    {
        Assert.NotEmpty(WordStyles.KnownBuiltInStyles);

        foreach (var entry in WordStyles.KnownBuiltInStyles)
        {
            Assert.Equal(entry.Value, WordStyles.Resolve(entry.Key));
        }
    }
}
