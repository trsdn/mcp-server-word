using System.IO.Compression;
using System.Text;

namespace WordMcp.ComInterop;

/// <summary>
/// Writes a minimal, valid empty Word document without involving Word itself.
/// </summary>
/// <remarks>
/// Creating documents through <c>Documents.Add</c> followed by <c>SaveAs2</c> is unreliable on
/// machines signed in to Microsoft 365: Word's AutoSave claims the new document for OneDrive and
/// silently ignores the requested local path, after which every save attempt either fails or blocks
/// on a modal dialog. Writing the package ourselves and opening it as an existing file avoids that
/// entirely and is also an order of magnitude faster.
/// </remarks>
public static class EmptyDocumentFactory
{
    private const string DocumentContentType =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";

    private const string MacroEnabledDocumentContentType =
        "application/vnd.ms-word.document.macroEnabled.main+xml";

    private const string RelsXml =
        """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>
        """;

    private const string DocumentXml =
        """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>
        """;

    /// <summary>
    /// Creates an empty document at the given path, overwriting an existing file.
    /// </summary>
    /// <param name="path">Absolute path of the document to create.</param>
    /// <param name="isMacroEnabled">Whether to write a macro-enabled (.docm) package.</param>
    public static void Create(string path, bool isMacroEnabled)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);

        using var stream = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Create);

        WriteEntry(archive, "[Content_Types].xml", BuildContentTypes(isMacroEnabled));
        WriteEntry(archive, "_rels/.rels", RelsXml);
        WriteEntry(archive, "word/document.xml", DocumentXml);
    }

    private static string BuildContentTypes(bool isMacroEnabled)
    {
        string documentContentType = isMacroEnabled
            ? MacroEnabledDocumentContentType
            : DocumentContentType;

        return $"""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="{documentContentType}"/></Types>
            """;
    }

    private static void WriteEntry(ZipArchive archive, string entryName, string content)
    {
        ZipArchiveEntry entry = archive.CreateEntry(entryName, CompressionLevel.Optimal);
        using Stream entryStream = entry.Open();
        using var writer = new StreamWriter(entryStream, new UTF8Encoding(false));
        writer.Write(content);
    }
}
