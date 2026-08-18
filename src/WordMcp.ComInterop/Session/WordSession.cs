using System.Diagnostics.CodeAnalysis;
using Microsoft.Extensions.Logging;

namespace WordMcp.ComInterop.Session;

/// <summary>
/// Entry point for Word COM interop. Every batch runs on its own STA thread with proper COM cleanup.
/// </summary>
public static class WordSession
{
    /// <summary>
    /// Begins a batch against one or more existing Word documents. The first path is the primary document.
    /// </summary>
    /// <param name="filePaths">Paths to existing Word documents.</param>
    /// <returns>A batch for executing operations.</returns>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility", Justification = "Windows-only COM automation project")]
    public static IWordBatch BeginBatch(params string[] filePaths)
        => BeginBatch(visible: false, operationTimeout: null, logger: null, filePaths);

    /// <summary>
    /// Begins a batch against one or more existing Word documents.
    /// </summary>
    /// <param name="visible">Whether the Word window is shown (default: hidden background automation).</param>
    /// <param name="operationTimeout">Maximum duration of a single operation (default: 5 minutes).</param>
    /// <param name="logger">Optional logger for diagnostics.</param>
    /// <param name="filePaths">Paths to existing Word documents. The first path is the primary document.</param>
    /// <returns>A batch for executing operations.</returns>
    /// <exception cref="ArgumentException">No path was supplied, or an extension is unsupported.</exception>
    /// <exception cref="FileNotFoundException">A document does not exist.</exception>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility", Justification = "Windows-only COM automation project")]
    public static IWordBatch BeginBatch(
        bool visible,
        TimeSpan? operationTimeout,
        ILogger? logger,
        params string[] filePaths)
    {
        if (filePaths == null || filePaths.Length == 0)
            throw new ArgumentException("At least one file path is required", nameof(filePaths));

        var fullPaths = new string[filePaths.Length];
        for (int i = 0; i < filePaths.Length; i++)
        {
            string fullPath = Path.GetFullPath(filePaths[i]);

            if (!File.Exists(fullPath))
            {
                throw new FileNotFoundException(
                    $"Word document not found: {fullPath}. To create a new document use the 'create' action instead of 'open'.",
                    fullPath);
            }

            FileAccessValidator.ValidateExtension(fullPath);
            fullPaths[i] = fullPath;
        }

        return new WordBatch(fullPaths, logger, visible, operationTimeout);
    }

    /// <summary>
    /// Creates a new Word document and returns an open batch for it.
    /// </summary>
    /// <param name="filePath">Path of the document to create. Its directory must exist.</param>
    /// <param name="visible">Whether the Word window is shown.</param>
    /// <param name="operationTimeout">Maximum duration of a single operation.</param>
    /// <param name="logger">Optional logger for diagnostics.</param>
    /// <returns>A batch with the new document open.</returns>
    /// <exception cref="ArgumentException">The extension is not a supported Word format.</exception>
    /// <exception cref="IOException">The file already exists.</exception>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility", Justification = "Windows-only COM automation project")]
    public static IWordBatch CreateNew(
        string filePath,
        bool visible = false,
        TimeSpan? operationTimeout = null,
        ILogger? logger = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(filePath);

        string fullPath = Path.GetFullPath(filePath);
        FileAccessValidator.ValidateExtension(fullPath);

        if (File.Exists(fullPath))
        {
            throw new IOException(
                $"File already exists: {fullPath}. Use the 'open' action to work with an existing document.");
        }

        string? directory = Path.GetDirectoryName(fullPath);
        if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
        {
            Directory.CreateDirectory(directory);
        }

        bool isMacroEnabled = ComInteropConstants.MacroEnabledExtensions
            .Contains(Path.GetExtension(fullPath).ToLowerInvariant());

        return WordBatch.CreateNewDocument(fullPath, isMacroEnabled, logger, visible, operationTimeout);
    }
}
