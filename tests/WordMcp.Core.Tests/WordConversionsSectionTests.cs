using WordMcp.Core.Utilities;
using Xunit;

namespace WordMcp.Core.Tests;

/// <summary>
/// Conversions introduced for the section and header/footer tools.
/// </summary>
public class WordConversionsSectionTests
{
    [Theory]
    [InlineData("portrait", 0)]
    [InlineData("Portrait", 0)]
    [InlineData("landscape", 1)]
    [InlineData("horizontal", 1)]
    public void ToWdOrientation_MapsKnownNames(string name, int expected)
        => Assert.Equal(expected, WordConversions.ToWdOrientation(name));

    [Theory]
    [InlineData("diagonal")]
    [InlineData(" ")]
    public void ToWdOrientation_RejectsUnknownNames(string name)
        => Assert.Throws<ArgumentException>(() => WordConversions.ToWdOrientation(name));

    [Fact]
    public void FromWdOrientation_RoundTrips()
    {
        Assert.Equal("portrait", WordConversions.FromWdOrientation(0));
        Assert.Equal("landscape", WordConversions.FromWdOrientation(1));
        Assert.Equal("other", WordConversions.FromWdOrientation(99));
    }

    [Theory]
    [InlineData("a4", 7)]
    [InlineData("A4", 7)]
    [InlineData("letter", 2)]
    [InlineData("legal", 5)]
    [InlineData("tabloid", 16)]
    public void ToWdPaperSize_MapsKnownNames(string name, int expected)
        => Assert.Equal(expected, WordConversions.ToWdPaperSize(name));

    [Fact]
    public void ToWdPaperSize_RejectsUnknownNames()
    {
        var ex = Assert.Throws<ArgumentException>(() => WordConversions.ToWdPaperSize("a2"));
        Assert.Contains("a4", ex.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("next-page", 2)]
    [InlineData("next_page", 2)]
    [InlineData("continuous", 0)]
    [InlineData("even-page", 3)]
    [InlineData("odd-page", 4)]
    [InlineData("new-column", 1)]
    public void ToWdSectionStart_MapsKnownNames(string name, int expected)
        => Assert.Equal(expected, WordConversions.ToWdSectionStart(name));

    [Theory]
    [InlineData("next-page", 2)]
    [InlineData("continuous", 3)]
    [InlineData("even-page", 4)]
    [InlineData("odd-page", 5)]
    public void ToWdSectionBreak_MapsKnownNames(string name, int expected)
        => Assert.Equal(expected, WordConversions.ToWdSectionBreak(name));

    [Fact]
    public void ToWdSectionBreak_RejectsNewColumn()
    {
        // A new column is a section start but not a section break type.
        Assert.Throws<ArgumentException>(() => WordConversions.ToWdSectionBreak("new-column"));
    }

    [Fact]
    public void FromWdSectionStart_RoundTrips()
    {
        foreach (string name in new[] { "next-page", "continuous", "even-page", "odd-page", "new-column" })
        {
            Assert.Equal(name, WordConversions.FromWdSectionStart(WordConversions.ToWdSectionStart(name)));
        }
    }

    [Theory]
    [InlineData("primary", 1)]
    [InlineData("default", 1)]
    [InlineData("first-page", 2)]
    [InlineData("first_page", 2)]
    [InlineData("even-pages", 3)]
    public void ToWdHeaderFooterIndex_MapsKnownNames(string name, int expected)
        => Assert.Equal(expected, WordConversions.ToWdHeaderFooterIndex(name));

    [Fact]
    public void ToWdHeaderFooterIndex_RejectsUnknownNames()
        => Assert.Throws<ArgumentException>(() => WordConversions.ToWdHeaderFooterIndex("last-page"));

    [Fact]
    public void FromWdHeaderFooterIndex_RoundTrips()
    {
        foreach (string name in new[] { "primary", "first-page", "even-pages" })
        {
            Assert.Equal(name, WordConversions.FromWdHeaderFooterIndex(
                WordConversions.ToWdHeaderFooterIndex(name)));
        }
    }
}
