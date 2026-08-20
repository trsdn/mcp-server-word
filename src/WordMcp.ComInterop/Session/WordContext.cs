using Word = Microsoft.Office.Interop.Word;

namespace WordMcp.ComInterop.Session;

/// <summary>
/// Provides access to the Word COM objects an operation runs against.
/// </summary>
public sealed class WordContext
{
    /// <summary>
    /// Creates a new <see cref="WordContext"/>.
    /// </summary>
    /// <param name="documentPath">Full path of the primary document.</param>
    /// <param name="app">The Word.Application COM object.</param>
    /// <param name="document">The Word.Document COM object.</param>
    public WordContext(string documentPath, Word.Application app, Word.Document document)
    {
        DocumentPath = documentPath ?? throw new ArgumentNullException(nameof(documentPath));
        App = app ?? throw new ArgumentNullException(nameof(app));
        Document = document ?? throw new ArgumentNullException(nameof(document));
    }

    /// <summary>
    /// Gets the full path of the primary document.
    /// </summary>
    public string DocumentPath { get; }

    /// <summary>
    /// Gets the Word.Application COM object.
    /// </summary>
    public Word.Application App { get; }

    /// <summary>
    /// Gets the Word.Document COM object.
    /// </summary>
    public Word.Document Document { get; }
}
