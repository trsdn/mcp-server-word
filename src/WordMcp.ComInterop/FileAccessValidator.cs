namespace WordMcp.ComInterop;

/// <summary>
/// Validates file access and lock status before Word COM operations.
/// </summary>
public static class FileAccessValidator
{
    // OLE2 Compound Document Format signature. IRM/AIP-protected Word files are stored as OLE2
    // containers with an EncryptedPackage stream instead of the ZIP-based Office Open XML format.
    private static ReadOnlySpan<byte> Ole2Signature => [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];

    /// <summary>
    /// Detects whether the file is IRM/AIP-protected by checking for the OLE2 compound document
    /// signature. Note that legacy .doc files share this signature.
    /// </summary>
    /// <param name="filePath">The file path to inspect.</param>
    /// <returns><c>true</c> when the file starts with the OLE2 header; otherwise <c>false</c>.</returns>
    public static bool IsOle2Container(string filePath)
    {
        if (!File.Exists(filePath))
            return false;

        try
        {
            Span<byte> header = stackalloc byte[8];
            using var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            int read = fs.Read(header);
            if (read < 8)
                return false;
            return header.SequenceEqual(Ole2Signature);
        }
        catch (IOException)
        {
            return false;
        }
        catch (UnauthorizedAccessException)
        {
            return false;
        }
    }

    /// <summary>
    /// Detects whether an Open XML Word file (.docx/.docm) is IRM/AIP-encrypted.
    /// Encrypted Open XML files are wrapped in an OLE2 container instead of a ZIP archive.
    /// </summary>
    /// <param name="filePath">The file path to inspect.</param>
    /// <returns><c>true</c> when the file is an encrypted Open XML document.</returns>
    public static bool IsIrmProtected(string filePath)
    {
        var extension = Path.GetExtension(filePath).ToLowerInvariant();

        // Legacy .doc/.dot files are legitimately OLE2 containers — not an IRM signal.
        if (extension is ".doc" or ".dot")
            return false;

        return IsOle2Container(filePath);
    }

    /// <summary>
    /// Validates that a file is not locked by attempting to open it with exclusive access.
    /// This is a fast OS-level check that does not require launching Word.
    /// </summary>
    /// <param name="filePath">The file path to validate.</param>
    /// <exception cref="InvalidOperationException">The file is locked or inaccessible.</exception>
    public static void ValidateFileNotLocked(string filePath)
    {
        try
        {
            using var lockTest = new FileStream(
                filePath,
                FileMode.Open,
                FileAccess.ReadWrite,
                FileShare.None);
            // Not locked — close and proceed.
        }
        catch (IOException ioEx)
        {
            throw CreateFileLockedError(filePath, ioEx);
        }
        catch (UnauthorizedAccessException uaEx)
        {
            throw new InvalidOperationException(
                $"Cannot access '{Path.GetFileName(filePath)}'. " +
                "The file may be read-only, you may lack permissions, or it is locked by another process. " +
                "Verify file permissions and close any application using this file.",
                uaEx);
        }
    }

    /// <summary>
    /// Creates a standardized exception for file-locked scenarios.
    /// </summary>
    /// <param name="filePath">The locked file path.</param>
    /// <param name="innerException">The underlying exception.</param>
    /// <returns>An <see cref="InvalidOperationException"/> with user-facing guidance.</returns>
    public static InvalidOperationException CreateFileLockedError(string filePath, Exception innerException)
    {
        return new InvalidOperationException(
            $"Cannot open '{Path.GetFileName(filePath)}'. " +
            "The file is already open in Word or another process is using it. " +
            "Close the file before running automation commands — WordMcp requires exclusive access.",
            innerException);
    }

    /// <summary>
    /// Validates that the extension of the given path is a Word format supported by this server.
    /// </summary>
    /// <param name="filePath">The file path to check.</param>
    /// <exception cref="ArgumentException">The extension is not a supported Word format.</exception>
    public static void ValidateExtension(string filePath)
    {
        var extension = Path.GetExtension(filePath).ToLowerInvariant();
        if (!ComInteropConstants.SupportedExtensions.Contains(extension))
        {
            throw new ArgumentException(
                $"Invalid file extension '{extension}'. Supported Word formats: " +
                string.Join(", ", ComInteropConstants.SupportedExtensions),
                nameof(filePath));
        }
    }
}
