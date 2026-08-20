namespace WordMcp.Core.Models;

/// <summary>
/// Metadata for a table in the document.
/// </summary>
public sealed class TableInfo
{
    /// <summary>Gets or sets the 1-based table index.</summary>
    public int Index { get; set; }

    /// <summary>Gets or sets the number of rows.</summary>
    public int RowCount { get; set; }

    /// <summary>Gets or sets the number of columns.</summary>
    public int ColumnCount { get; set; }

    /// <summary>Gets or sets the applied table style name.</summary>
    public string Style { get; set; } = string.Empty;
}

/// <summary>
/// All tables of a document.
/// </summary>
public sealed class TableListResult : ResultBase
{
    /// <summary>Gets or sets the total number of tables.</summary>
    public int TotalCount { get; set; }

    /// <summary>Gets or sets the tables.</summary>
    public IReadOnlyList<TableInfo> Tables { get; set; } = [];
}

/// <summary>
/// Cell contents of a table.
/// </summary>
public sealed class TableReadResult : ResultBase
{
    /// <summary>Gets or sets the table metadata.</summary>
    public TableInfo? Table { get; set; }

    /// <summary>Gets or sets the cell values as rows of columns.</summary>
    public IReadOnlyList<IReadOnlyList<string>> Rows { get; set; } = [];
}

/// <summary>
/// Result of an operation that modifies a table.
/// </summary>
public sealed class TableResult : ResultBase
{
    /// <summary>Gets or sets the affected table.</summary>
    public TableInfo? Table { get; set; }
}
