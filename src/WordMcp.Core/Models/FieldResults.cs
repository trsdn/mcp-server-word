namespace WordMcp.Core.Models;

/// <summary>
/// Metadata for a field in the document.
/// </summary>
public sealed class FieldInfo
{
    /// <summary>Gets or sets the 1-based field index.</summary>
    public int Index { get; set; }

    /// <summary>Gets or sets the numeric <c>WdFieldType</c> value.</summary>
    public int Type { get; set; }

    /// <summary>Gets or sets the field code, for example <c>TOC \o "1-3"</c>.</summary>
    public string Code { get; set; } = string.Empty;

    /// <summary>Gets or sets the last calculated field result.</summary>
    public string Result { get; set; } = string.Empty;
}

/// <summary>
/// All fields of a document.
/// </summary>
public sealed class FieldListResult : ResultBase
{
    /// <summary>Gets or sets the total number of fields in the main text.</summary>
    public int TotalCount { get; set; }

    /// <summary>Gets or sets the fields.</summary>
    public IReadOnlyList<FieldInfo> Fields { get; set; } = [];
}

/// <summary>
/// Result of an operation that inserts or updates fields.
/// </summary>
public sealed class FieldResult : ResultBase
{
    /// <summary>Gets or sets the number of fields affected by the operation.</summary>
    public int UpdatedCount { get; set; }

    /// <summary>
    /// Gets or sets the number of entries a table of contents contains after updating.
    /// </summary>
    /// <remarks>
    /// Zero means the document has no paragraphs with heading styles, so the table of contents
    /// is empty even though it was inserted successfully.
    /// </remarks>
    public int? EntryCount { get; set; }
}
