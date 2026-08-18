using WordMcp.ComInterop;
using WordMcp.ComInterop.Session;
using WordMcp.Core.Models;
using WordMcp.Core.Utilities;

namespace WordMcp.Core.Commands.Table;

/// <summary>
/// Word COM implementation of <see cref="ITableCommands"/>.
/// </summary>
public sealed class TableCommands : ITableCommands
{
    /// <inheritdoc />
    public TableListResult List(IWordBatch batch)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            dynamic tables = ctx.Document.Tables;
            int total = (int)tables.Count;

            var list = new List<TableInfo>(total);
            for (int i = 1; i <= total; i++)
            {
                ct.ThrowIfCancellationRequested();
                list.Add(Describe(tables[i], i));
            }

            return new TableListResult
            {
                TotalCount = total,
                Tables = list
            };
        });
    }

    /// <inheritdoc />
    public TableResult Create(IWordBatch batch, int rows, int columns, string? style = null)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentOutOfRangeException.ThrowIfLessThan(rows, 1);
        ArgumentOutOfRangeException.ThrowIfLessThan(columns, 1);

        return batch.Execute((ctx, ct) =>
        {
            dynamic doc = ctx.Document;
            dynamic range = doc.Content;
            range.Collapse(ComInteropConstants.WdCollapseEnd);
            range.InsertParagraphAfter();

            dynamic target = doc.Content;
            target.Collapse(ComInteropConstants.WdCollapseEnd);

            dynamic table = doc.Tables.Add(target, rows, columns);
            ApplyStyle(table, style ?? "Table Grid");

            int index = (int)doc.Tables.Count;

            return new TableResult
            {
                Table = Describe(table, index),
                Message = $"Table {index} created with {rows} row(s) and {columns} column(s)."
            };
        });
    }

    /// <inheritdoc />
    public TableReadResult Read(IWordBatch batch, int index)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentOutOfRangeException.ThrowIfLessThan(index, 1);

        return batch.Execute((ctx, ct) =>
        {
            dynamic table = GetTable(ctx.Document, index);

            int rowCount = (int)table.Rows.Count;
            int columnCount = (int)table.Columns.Count;

            var rows = new List<IReadOnlyList<string>>(rowCount);
            for (int r = 1; r <= rowCount; r++)
            {
                ct.ThrowIfCancellationRequested();

                var cells = new List<string>(columnCount);
                for (int c = 1; c <= columnCount; c++)
                {
                    cells.Add(ReadCell(table, r, c));
                }

                rows.Add(cells);
            }

            return new TableReadResult
            {
                Table = Describe(table, index),
                Rows = rows
            };
        });
    }

    /// <inheritdoc />
    public OperationResult SetCell(IWordBatch batch, int index, int row, int column, string text)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentNullException.ThrowIfNull(text);
        ArgumentOutOfRangeException.ThrowIfLessThan(index, 1);
        ArgumentOutOfRangeException.ThrowIfLessThan(row, 1);
        ArgumentOutOfRangeException.ThrowIfLessThan(column, 1);

        batch.Execute((ctx, ct) =>
        {
            dynamic table = GetTable(ctx.Document, index);
            table.Cell(row, column).Range.Text = text;
        });

        return OperationResult.Ok($"Cell ({row}, {column}) of table {index} updated.");
    }

    /// <inheritdoc />
    public TableResult AddRow(IWordBatch batch, int index, IReadOnlyList<string>? values = null)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentOutOfRangeException.ThrowIfLessThan(index, 1);

        return batch.Execute((ctx, ct) =>
        {
            dynamic table = GetTable(ctx.Document, index);
            table.Rows.Add();

            int rowNumber = (int)table.Rows.Count;

            if (values is { Count: > 0 })
            {
                int columnCount = (int)table.Columns.Count;
                for (int c = 1; c <= Math.Min(columnCount, values.Count); c++)
                {
                    table.Cell(rowNumber, c).Range.Text = values[c - 1];
                }
            }

            return new TableResult
            {
                Table = Describe(table, index),
                Message = $"Row {rowNumber} added to table {index}."
            };
        });
    }

    /// <inheritdoc />
    public TableResult DeleteRow(IWordBatch batch, int index, int row)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentOutOfRangeException.ThrowIfLessThan(index, 1);
        ArgumentOutOfRangeException.ThrowIfLessThan(row, 1);

        return batch.Execute((ctx, ct) =>
        {
            dynamic table = GetTable(ctx.Document, index);

            int rowCount = (int)table.Rows.Count;
            if (row > rowCount)
            {
                throw new ArgumentOutOfRangeException(
                    nameof(row), $"Row {row} does not exist. Table {index} has {rowCount} row(s).");
            }

            table.Rows[row].Delete();

            return new TableResult
            {
                Table = Describe(table, index),
                Message = $"Row {row} deleted from table {index}."
            };
        });
    }

    /// <inheritdoc />
    public TableResult SetStyle(IWordBatch batch, int index, string style)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentException.ThrowIfNullOrWhiteSpace(style);
        ArgumentOutOfRangeException.ThrowIfLessThan(index, 1);

        return batch.Execute((ctx, ct) =>
        {
            dynamic table = GetTable(ctx.Document, index);
            ApplyStyle(table, style);

            return new TableResult
            {
                Table = Describe(table, index),
                Message = $"Style '{style}' applied to table {index}."
            };
        });
    }

    private static dynamic GetTable(dynamic document, int index)
    {
        int total = (int)document.Tables.Count;
        if (index > total)
        {
            throw new ArgumentOutOfRangeException(
                nameof(index), $"Table {index} does not exist. The document has {total} table(s).");
        }

        return document.Tables[index];
    }

    private static string ReadCell(dynamic table, int row, int column)
    {
        try
        {
            return WordConversions.CleanRangeText((string?)table.Cell(row, column).Range.Text);
        }
        catch (System.Runtime.InteropServices.COMException)
        {
            // Merged cells raise "the requested member of the collection does not exist".
            return string.Empty;
        }
    }

    private static void ApplyStyle(dynamic table, string? style)
    {
        if (string.IsNullOrWhiteSpace(style))
            return;

        try
        {
            table.Style = style;
        }
        catch (System.Runtime.InteropServices.COMException ex)
        {
            throw new ArgumentException(
                $"Table style '{style}' does not exist in this document.", nameof(style), ex);
        }
    }

    private static TableInfo Describe(dynamic table, int index)
    {
        string styleName;
        try
        {
            dynamic style = table.Style;
            styleName = (string)style.NameLocal;
        }
        catch (System.Runtime.InteropServices.COMException)
        {
            styleName = string.Empty;
        }

        return new TableInfo
        {
            Index = index,
            RowCount = (int)table.Rows.Count,
            ColumnCount = (int)table.Columns.Count,
            Style = styleName
        };
    }
}
