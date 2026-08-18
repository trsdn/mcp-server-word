namespace WordMcp.Core.Models;

/// <summary>
/// Base class for every command result. Serialized to JSON and returned to the MCP client.
/// </summary>
public abstract class ResultBase
{
    /// <summary>Gets or sets a value indicating whether the operation succeeded.</summary>
    public bool Success { get; set; } = true;

    /// <summary>Gets or sets an optional human-readable message.</summary>
    public string? Message { get; set; }
}

/// <summary>
/// Generic result for operations that only report success or failure.
/// </summary>
public sealed class OperationResult : ResultBase
{
    /// <summary>
    /// Creates a successful result.
    /// </summary>
    /// <param name="message">Optional message describing what happened.</param>
    /// <returns>A successful <see cref="OperationResult"/>.</returns>
    public static OperationResult Ok(string? message = null)
        => new() { Success = true, Message = message };
}

/// <summary>
/// Document statistics and metadata.
/// </summary>
public sealed class DocumentInfoResult : ResultBase
{
    /// <summary>Gets or sets the full path of the document.</summary>
    public string FilePath { get; set; } = string.Empty;

    /// <summary>Gets or sets the file name of the document.</summary>
    public string FileName { get; set; } = string.Empty;

    /// <summary>Gets or sets the number of words.</summary>
    public int WordCount { get; set; }

    /// <summary>Gets or sets the number of characters.</summary>
    public int CharacterCount { get; set; }

    /// <summary>Gets or sets the number of paragraphs.</summary>
    public int ParagraphCount { get; set; }

    /// <summary>Gets or sets the number of pages.</summary>
    public int PageCount { get; set; }

    /// <summary>Gets or sets the number of tables.</summary>
    public int TableCount { get; set; }

    /// <summary>Gets or sets the number of inline shapes (images).</summary>
    public int InlineShapeCount { get; set; }

    /// <summary>Gets or sets the number of sections.</summary>
    public int SectionCount { get; set; }

    /// <summary>Gets or sets a value indicating whether the document has unsaved changes.</summary>
    public bool HasUnsavedChanges { get; set; }

    /// <summary>Gets or sets a value indicating whether the document is read-only.</summary>
    public bool IsReadOnly { get; set; }
}

/// <summary>
/// Built-in document properties.
/// </summary>
public sealed class DocumentPropertiesResult : ResultBase
{
    /// <summary>Gets or sets the document title.</summary>
    public string? Title { get; set; }

    /// <summary>Gets or sets the author.</summary>
    public string? Author { get; set; }

    /// <summary>Gets or sets the subject.</summary>
    public string? Subject { get; set; }

    /// <summary>Gets or sets the keywords.</summary>
    public string? Keywords { get; set; }

    /// <summary>Gets or sets the comments.</summary>
    public string? Comments { get; set; }

    /// <summary>Gets or sets the company.</summary>
    public string? Company { get; set; }

    /// <summary>Gets or sets the last author.</summary>
    public string? LastAuthor { get; set; }
}

/// <summary>
/// Result of an export or save-as operation.
/// </summary>
public sealed class ExportResult : ResultBase
{
    /// <summary>Gets or sets the path of the produced file.</summary>
    public string OutputPath { get; set; } = string.Empty;

    /// <summary>Gets or sets the size of the produced file in bytes.</summary>
    public long FileSizeBytes { get; set; }
}
