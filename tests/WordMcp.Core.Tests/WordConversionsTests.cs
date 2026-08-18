using WordMcp.ComInterop;
using WordMcp.Core.Utilities;
using Xunit;

namespace WordMcp.Core.Tests;

public class WordConversionsTests
{
    [Theory]
    [InlineData("left", ComInteropConstants.WdAlignParagraphLeft)]
    [InlineData("LEFT", ComInteropConstants.WdAlignParagraphLeft)]
    [InlineData("center", ComInteropConstants.WdAlignParagraphCenter)]
    [InlineData("centre", ComInteropConstants.WdAlignParagraphCenter)]
    [InlineData("centered", ComInteropConstants.WdAlignParagraphCenter)]
    [InlineData("right", ComInteropConstants.WdAlignParagraphRight)]
    [InlineData(" justify ", ComInteropConstants.WdAlignParagraphJustify)]
    [InlineData("justified", ComInteropConstants.WdAlignParagraphJustify)]
    public void ToWdAlignment_MapsKnownNames(string input, int expected)
        => Assert.Equal(expected, WordConversions.ToWdAlignment(input));

    [Theory]
    [InlineData("middle")]
    [InlineData("top")]
    public void ToWdAlignment_RejectsUnknownNames(string input)
        => Assert.Throws<ArgumentException>(() => WordConversions.ToWdAlignment(input));

    [Fact]
    public void ToWdAlignment_RejectsEmptyInput()
        => Assert.Throws<ArgumentException>(() => WordConversions.ToWdAlignment("  "));

    [Theory]
    [InlineData(ComInteropConstants.WdAlignParagraphLeft, "left")]
    [InlineData(ComInteropConstants.WdAlignParagraphCenter, "center")]
    [InlineData(ComInteropConstants.WdAlignParagraphRight, "right")]
    [InlineData(ComInteropConstants.WdAlignParagraphJustify, "justify")]
    [InlineData(99, "other")]
    public void FromWdAlignment_MapsBack(int input, string expected)
        => Assert.Equal(expected, WordConversions.FromWdAlignment(input));

    [Theory]
    [InlineData("left")]
    [InlineData("center")]
    [InlineData("right")]
    [InlineData("justify")]
    public void Alignment_RoundTrips(string alignment)
        => Assert.Equal(alignment, WordConversions.FromWdAlignment(WordConversions.ToWdAlignment(alignment)));

    [Theory]
    [InlineData(@"C:\docs\a.docx", ComInteropConstants.WdFormatXmlDocument)]
    [InlineData(@"C:\docs\a.DOCX", ComInteropConstants.WdFormatXmlDocument)]
    [InlineData(@"C:\docs\a.docm", ComInteropConstants.WdFormatXmlDocumentMacroEnabled)]
    [InlineData(@"C:\docs\a.doc", ComInteropConstants.WdFormatDocument97)]
    [InlineData(@"C:\docs\a.pdf", ComInteropConstants.WdFormatPdf)]
    [InlineData(@"C:\docs\a.rtf", ComInteropConstants.WdFormatRtf)]
    [InlineData(@"C:\docs\a.txt", ComInteropConstants.WdFormatText)]
    [InlineData(@"C:\docs\a.html", ComInteropConstants.WdFormatFilteredHtml)]
    [InlineData(@"C:\docs\a.htm", ComInteropConstants.WdFormatFilteredHtml)]
    public void ToWdSaveFormat_MapsSupportedExtensions(string path, int expected)
        => Assert.Equal(expected, WordConversions.ToWdSaveFormat(path));

    [Theory]
    [InlineData(@"C:\docs\a.pptx")]
    [InlineData(@"C:\docs\a")]
    public void ToWdSaveFormat_RejectsUnsupportedExtensions(string path)
        => Assert.Throws<ArgumentException>(() => WordConversions.ToWdSaveFormat(path));

    [Theory]
    [InlineData("Hello\r", "Hello")]
    [InlineData("Cell\r\a", "Cell")]
    [InlineData("Cell\a", "Cell")]
    [InlineData("Plain", "Plain")]
    [InlineData("", "")]
    [InlineData(null, "")]
    public void CleanRangeText_StripsWordControlCharacters(string? input, string expected)
        => Assert.Equal(expected, WordConversions.CleanRangeText(input));

    [Fact]
    public void CleanRangeText_KeepsInnerCarriageReturns()
        => Assert.Equal("a\rb", WordConversions.CleanRangeText("a\rb\r"));

    [Theory]
    // Word stores colours as BGR, so the red and blue bytes swap.
    [InlineData("#0078D4", 0xD47800)]
    [InlineData("0078D4", 0xD47800)]
    [InlineData("#FF0000", 0x0000FF)]
    [InlineData("#00FF00", 0x00FF00)]
    [InlineData("#0000FF", 0xFF0000)]
    [InlineData("#000000", 0x000000)]
    [InlineData("#FFFFFF", 0xFFFFFF)]
    public void ToWdColor_ConvertsRgbToBgr(string hex, int expected)
        => Assert.Equal(expected, WordConversions.ToWdColor(hex));

    [Theory]
    [InlineData("#12345")]
    [InlineData("#1234567")]
    [InlineData("#GGGGGG")]
    [InlineData("blue")]
    public void ToWdColor_RejectsInvalidValues(string hex)
        => Assert.Throws<ArgumentException>(() => WordConversions.ToWdColor(hex));
}
