using WordMcp.ComInterop;
using Xunit;

namespace WordMcp.Core.Tests;

public sealed class FileAccessValidatorTests : IDisposable
{
    private static ReadOnlySpan<byte> Ole2Signature => [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];

    private readonly string _directory =
        Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "wordmcp-tests-" + Guid.NewGuid().ToString("N"))).FullName;

    public void Dispose()
    {
        try
        {
            Directory.Delete(_directory, recursive: true);
        }
        catch (IOException)
        {
            // Best effort cleanup: a leftover temp directory must not fail the test run.
        }
    }

    private string WriteFile(string name, ReadOnlySpan<byte> content)
    {
        var path = Path.Combine(_directory, name);
        File.WriteAllBytes(path, content.ToArray());
        return path;
    }

    [Fact]
    public void IsOle2Container_DetectsSignature()
    {
        var path = WriteFile("legacy.doc", Ole2Signature);
        Assert.True(FileAccessValidator.IsOle2Container(path));
    }

    [Fact]
    public void IsOle2Container_ReturnsFalseForZipBasedOpenXml()
    {
        // Open XML files are ZIP archives, which start with "PK\x03\x04".
        var path = WriteFile("modern.docx", [0x50, 0x4B, 0x03, 0x04, 0x00, 0x00, 0x00, 0x00]);
        Assert.False(FileAccessValidator.IsOle2Container(path));
    }

    [Fact]
    public void IsOle2Container_ReturnsFalseForTruncatedFile()
    {
        var path = WriteFile("short.docx", [0xD0, 0xCF]);
        Assert.False(FileAccessValidator.IsOle2Container(path));
    }

    [Fact]
    public void IsOle2Container_ReturnsFalseForMissingFile()
        => Assert.False(FileAccessValidator.IsOle2Container(Path.Combine(_directory, "nope.docx")));

    [Theory]
    [InlineData("legacy.doc")]
    [InlineData("legacy.dot")]
    [InlineData("LEGACY.DOC")]
    public void IsIrmProtected_IgnoresLegacyBinaryFormats(string name)
    {
        // Legacy binary Word files are legitimately OLE2 containers and must not be
        // misreported as rights-protected.
        var path = WriteFile(name, Ole2Signature);
        Assert.False(FileAccessValidator.IsIrmProtected(path));
    }

    [Fact]
    public void IsIrmProtected_DetectsEncryptedOpenXml()
    {
        var path = WriteFile("protected.docx", Ole2Signature);
        Assert.True(FileAccessValidator.IsIrmProtected(path));
    }

    [Fact]
    public void IsIrmProtected_ReturnsFalseForRegularOpenXml()
    {
        var path = WriteFile("plain.docx", [0x50, 0x4B, 0x03, 0x04, 0x00, 0x00, 0x00, 0x00]);
        Assert.False(FileAccessValidator.IsIrmProtected(path));
    }

    [Fact]
    public void ValidateFileNotLocked_PassesForFreeFile()
    {
        var path = WriteFile("free.docx", [0x50, 0x4B, 0x03, 0x04]);
        FileAccessValidator.ValidateFileNotLocked(path);
    }

    [Fact]
    public void ValidateFileNotLocked_ThrowsWhileFileIsHeldOpen()
    {
        var path = WriteFile("busy.docx", [0x50, 0x4B, 0x03, 0x04]);

        using var hold = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);

        var ex = Assert.Throws<InvalidOperationException>(() => FileAccessValidator.ValidateFileNotLocked(path));
        Assert.Contains("busy.docx", ex.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(".docx")]
    [InlineData(".docm")]
    [InlineData(".doc")]
    [InlineData(".dotx")]
    [InlineData(".dotm")]
    [InlineData(".rtf")]
    public void ValidateExtension_AcceptsSupportedFormats(string extension)
        => FileAccessValidator.ValidateExtension(@"C:\docs\file" + extension);

    [Theory]
    [InlineData(".pptx")]
    [InlineData(".pdf")]
    [InlineData(".txt")]
    [InlineData("")]
    public void ValidateExtension_RejectsOtherFormats(string extension)
        => Assert.Throws<ArgumentException>(() => FileAccessValidator.ValidateExtension(@"C:\docs\file" + extension));

    [Fact]
    public void CreateFileLockedError_MentionsExclusiveAccess()
    {
        var ex = FileAccessValidator.CreateFileLockedError(@"C:\docs\report.docx", new IOException("locked"));

        Assert.Contains("report.docx", ex.Message, StringComparison.Ordinal);
        Assert.Contains("exclusive access", ex.Message, StringComparison.Ordinal);
        Assert.IsType<IOException>(ex.InnerException);
    }
}
