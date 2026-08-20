using WordMcp.ComInterop.Session;
using WordMcp.Core.Attributes;
using WordMcp.Core.Models;

namespace WordMcp.Core.Commands.Table;

/// <summary>
/// Table operations: listing, creating, reading and editing tables.
/// </summary>
[ServiceCategory("table", "Table")]
[McpTool("table",
    Title = "Table Operations",
    Description = "Table operations on an open document. "
        + "table(list, session_id) returns index, size and style of every table. "
        + "table(create, session_id, rows=3, columns=4, style='Table Grid') appends a table. "
        + "table(read, session_id, index=1) returns all cell values as rows. "
        + "table(set-cell, session_id, index=1, row=1, column=2, text='Total') writes one cell. "
        + "table(add-row, session_id, index=1, values=['a','b']) appends a filled row. "
        + "table(delete-row|set-style, session_id, index=1, ...) edits an existing table. "
        + "All indexes are 1-based.")]
public interface ITableCommands
{
    /// <summary>
    /// Lists all tables of the document.
    /// </summary>
    /// <param name="batch">The Word batch to operate on.</param>
    /// <returns>The tables of the document.</returns>
    [ServiceAction("list")]
    TableListResult List(IWordBatch batch);

    /// <summary>
    /// Appends a new table at the end of the document.
    /// </summary>
    /// <param name="batch">The Word batch to operate on.</param>
    /// <param name="rows">Number of rows.</param>
    /// <param name="columns">Number of columns.</param>
    /// <param name="style">Optional table style name, for example <c>Table Grid</c>.</param>
    /// <returns>The created table.</returns>
    [ServiceAction("create")]
    TableResult Create(IWordBatch batch, int rows, int columns, string? style = null);

    /// <summary>
    /// Reads the cell contents of a table.
    /// </summary>
    /// <param name="batch">The Word batch to operate on.</param>
    /// <param name="index">1-based table index.</param>
    /// <returns>The table contents.</returns>
    [ServiceAction("read")]
    TableReadResult Read(IWordBatch batch, int index);

    /// <summary>
    /// Writes the text of a single cell.
    /// </summary>
    /// <param name="batch">The Word batch to operate on.</param>
    /// <param name="index">1-based table index.</param>
    /// <param name="row">1-based row number.</param>
    /// <param name="column">1-based column number.</param>
    /// <param name="text">The cell text.</param>
    /// <returns>The result of the operation.</returns>
    [ServiceAction("set-cell")]
    OperationResult SetCell(IWordBatch batch, int index, int row, int column, string text = "");

    /// <summary>
    /// Appends a row to a table, optionally filling it with values.
    /// </summary>
    /// <param name="batch">The Word batch to operate on.</param>
    /// <param name="index">1-based table index.</param>
    /// <param name="values">Optional cell values, applied left to right.</param>
    /// <returns>The updated table.</returns>
    [ServiceAction("add-row")]
    TableResult AddRow(IWordBatch batch, int index, IReadOnlyList<string>? values = null);

    /// <summary>
    /// Deletes a row from a table.
    /// </summary>
    /// <param name="batch">The Word batch to operate on.</param>
    /// <param name="index">1-based table index.</param>
    /// <param name="row">1-based row number.</param>
    /// <returns>The updated table.</returns>
    [ServiceAction("delete-row")]
    TableResult DeleteRow(IWordBatch batch, int index, int row);

    /// <summary>
    /// Applies a table style.
    /// </summary>
    /// <param name="batch">The Word batch to operate on.</param>
    /// <param name="index">1-based table index.</param>
    /// <param name="style">Table style name, for example <c>Table Grid</c>.</param>
    /// <returns>The updated table.</returns>
    [ServiceAction("set-style")]
    TableResult SetStyle(IWordBatch batch, int index, string style);
}
