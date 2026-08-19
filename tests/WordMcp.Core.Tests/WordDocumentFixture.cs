using WordMcp.ComInterop.Session;
using Xunit;

namespace WordMcp.Core.Tests;

/// <summary>
/// Shared Word session for the integration tests. Starting Word is expensive, so all
/// integration tests share one instance and one document.
/// </summary>
public sealed class WordDocumentFixture : IDisposable
{
    public SessionManager Sessions { get; } = new();

    public string DirectoryPath { get; } =
        Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "wordmcp-it-" + Guid.NewGuid().ToString("N"))).FullName;

    public string SessionId { get; }

    public string FilePath { get; }

    public IWordBatch Batch => Sessions.GetBatch(SessionId);

    public WordDocumentFixture()
    {
        FilePath = Path.Combine(DirectoryPath, "integration.docx");
        SessionId = Sessions.Create(FilePath).SessionId;
    }

    public void Dispose()
    {
        Sessions.Dispose();

        try
        {
            Directory.Delete(DirectoryPath, recursive: true);
        }
        catch (IOException)
        {
            // Word may still hold a handle briefly; a leftover temp folder is harmless.
        }
        catch (UnauthorizedAccessException)
        {
            // Same as above.
        }
    }
}

/// <summary>
/// Marks the collection so that all Word integration tests run sequentially against one fixture.
/// </summary>
[CollectionDefinition(Name)]
public sealed class WordCollection : ICollectionFixture<WordDocumentFixture>
{
    public const string Name = "Word";
}
