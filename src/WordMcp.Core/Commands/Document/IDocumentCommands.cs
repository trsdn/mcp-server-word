using WordMcp.ComInterop.Session;
using WordMcp.Core.Attributes;
using WordMcp.Core.Models;

namespace WordMcp.Core.Commands.Document;

/// <summary>
/// Document-level operations: statistics, built-in properties and export.
/// </summary>
[ServiceCategory("document", "Document")]
[McpTool("document",
    Title = "Document Operations",
    Description = "Inspect document statistics, read and write document properties, export to PDF or other formats.")]
public interface IDocumentCommands
{
    /// <summary>
    /// Reads statistics and state of the open document.
    /// </summary>
    /// <param name="batch">The Word batch to operate on.</param>
    /// <returns>Document statistics.</returns>
    [ServiceAction("get-info")]
    DocumentInfoResult GetInfo(IWordBatch batch);

    /// <summary>
    /// Reads the built-in document properties.
    /// </summary>
    /// <param name="batch">The Word batch to operate on.</param>
    /// <returns>The document properties.</returns>
    [ServiceAction("get-properties")]
    DocumentPropertiesResult GetProperties(IWordBatch batch);

    /// <summary>
    /// Updates built-in document properties. Only non-null values are written.
    /// </summary>
    /// <param name="batch">The Word batch to operate on.</param>
    /// <param name="title">New title, or <c>null</c> to keep the current value.</param>
    /// <param name="author">New author, or <c>null</c> to keep the current value.</param>
    /// <param name="subject">New subject, or <c>null</c> to keep the current value.</param>
    /// <param name="keywords">New keywords, or <c>null</c> to keep the current value.</param>
    /// <param name="comments">New comments, or <c>null</c> to keep the current value.</param>
    /// <param name="company">New company, or <c>null</c> to keep the current value.</param>
    /// <returns>The properties after the update.</returns>
    [ServiceAction("set-properties")]
    DocumentPropertiesResult SetProperties(
        IWordBatch batch,
        string? title = null,
        string? author = null,
        string? subject = null,
        string? keywords = null,
        string? comments = null,
        string? company = null);

    /// <summary>
    /// Exports the document to PDF.
    /// </summary>
    /// <param name="batch">The Word batch to operate on.</param>
    /// <param name="outputPath">Target path ending in <c>.pdf</c>.</param>
    /// <returns>Details of the produced file.</returns>
    [ServiceAction("export-pdf")]
    ExportResult ExportPdf(IWordBatch batch, string outputPath);

    /// <summary>
    /// Saves a copy of the document in the format implied by the target extension.
    /// </summary>
    /// <param name="batch">The Word batch to operate on.</param>
    /// <param name="outputPath">Target path; the extension selects the format.</param>
    /// <returns>Details of the produced file.</returns>
    /// <remarks>
    /// For non-PDF targets this also persists the open document to its original path, because
    /// Word's object model offers no format-changing "save a copy" operation.
    /// </remarks>
    [ServiceAction("save-as")]
    ExportResult SaveAs(IWordBatch batch, string outputPath);
}
