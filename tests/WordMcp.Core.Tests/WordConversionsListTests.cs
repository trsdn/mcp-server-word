using WordMcp.ComInterop;
using WordMcp.Core.Utilities;
using Xunit;

namespace WordMcp.Core.Tests;

/// <summary>
/// Conversions between the friendly list names on the wire and Word's list constants.
/// </summary>
public class WordConversionsListTests
{
    [Theory]
    [InlineData("bullet", ComInteropConstants.WdBulletGallery)]
    [InlineData("bullets", ComInteropConstants.WdBulletGallery)]
    [InlineData("number", ComInteropConstants.WdNumberGallery)]
    [InlineData("numbered", ComInteropConstants.WdNumberGallery)]
    [InlineData("outline-number", ComInteropConstants.WdOutlineNumberGallery)]
    [InlineData("outline_number", ComInteropConstants.WdOutlineNumberGallery)]
    [InlineData("OUTLINE", ComInteropConstants.WdOutlineNumberGallery)]
    public void ToWdListGallery_MapsKnownNames(string listType, int expected)
        => Assert.Equal(expected, WordConversions.ToWdListGallery(listType));

    [Fact]
    public void ToWdListGallery_RejectsUnknownName()
        => Assert.Throws<ArgumentException>(() => WordConversions.ToWdListGallery("squiggle"));

    [Fact]
    public void ToWdListGallery_RejectsEmptyName()
        => Assert.Throws<ArgumentException>(() => WordConversions.ToWdListGallery("  "));

    [Theory]
    [InlineData(ComInteropConstants.WdListNoNumbering, "none")]
    [InlineData(ComInteropConstants.WdListBullet, "bullet")]
    [InlineData(ComInteropConstants.WdListSimpleNumbering, "number")]
    [InlineData(ComInteropConstants.WdListListNumOnly, "number")]
    [InlineData(ComInteropConstants.WdListOutlineNumbering, "outline-number")]
    [InlineData(ComInteropConstants.WdListMixedNumbering, "outline-number")]
    public void FromWdListType_MapsKnownConstants(int wdListType, string expected)
        => Assert.Equal(expected, WordConversions.FromWdListType(wdListType));

    [Fact]
    public void FromWdListType_FallsBackForUnknownConstants()
        => Assert.Equal("other", WordConversions.FromWdListType(99));
}
