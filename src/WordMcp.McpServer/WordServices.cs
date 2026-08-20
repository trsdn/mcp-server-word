using Microsoft.Extensions.Logging;
using WordMcp.ComInterop.Session;
using WordMcp.Core.Commands.Document;
using WordMcp.Core.Commands.Field;
using WordMcp.Core.Commands.Image;
using WordMcp.Core.Commands.Paragraph;
using WordMcp.Core.Commands.Table;
using WordMcp.Core.Commands.Text;

namespace WordMcp.McpServer;

/// <summary>
/// Process-wide singletons used by the MCP tools.
/// </summary>
/// <remarks>
/// The MCP SDK discovers tools as static methods, so the session manager and command
/// implementations are held here instead of being resolved through DI.
/// </remarks>
internal static class WordServices
{
    private static readonly Lock Gate = new();
    private static SessionManager? _sessions;

    /// <summary>Gets the shared session manager.</summary>
    public static SessionManager Sessions
    {
        get
        {
            lock (Gate)
            {
                return _sessions ??= new SessionManager(Logger);
            }
        }
    }

    /// <summary>Gets or sets the logger handed to new sessions.</summary>
    public static ILogger? Logger { get; set; }

    /// <summary>Gets the document command implementation.</summary>
    public static IDocumentCommands Documents { get; } = new DocumentCommands();

    /// <summary>Gets the text command implementation.</summary>
    public static ITextCommands Texts { get; } = new TextCommands();

    /// <summary>Gets the paragraph command implementation.</summary>
    public static IParagraphCommands Paragraphs { get; } = new ParagraphCommands();

    /// <summary>Gets the table command implementation.</summary>
    public static ITableCommands Tables { get; } = new TableCommands();

    /// <summary>Gets the image command implementation.</summary>
    public static IImageCommands Images { get; } = new ImageCommands();

    /// <summary>Gets the field command implementation.</summary>
    public static IFieldCommands Fields { get; } = new FieldCommands();

    /// <summary>
    /// Saves and closes every open session. Called on shutdown so unsaved work is not lost.
    /// </summary>
    public static void Shutdown()
    {
        lock (Gate)
        {
            _sessions?.Dispose();
            _sessions = null;
        }
    }
}
