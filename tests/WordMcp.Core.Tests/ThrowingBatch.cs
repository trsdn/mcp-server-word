using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;
using WordMcp.ComInterop.Session;

namespace WordMcp.Core.Tests;

/// <summary>
/// Stands in for a Word batch in validation tests. Any call means the argument checks let the
/// request through, which is what those tests assert.
/// </summary>
internal sealed class ThrowingBatch : IWordBatch
{
    public string DocumentPath => "fake.docx";

    public ILogger Logger => NullLogger.Instance;

    public TimeSpan OperationTimeout => TimeSpan.FromSeconds(30);

    public int? WordProcessId => null;

    public T Execute<T>(Func<WordContext, CancellationToken, T> operation, CancellationToken cancellationToken = default)
        => throw new NotSupportedException("Reached the batch.");

    public void Execute(Action<WordContext, CancellationToken> operation, CancellationToken cancellationToken = default)
        => throw new NotSupportedException("Reached the batch.");

    public void Save(CancellationToken cancellationToken = default)
        => throw new NotSupportedException("Reached the batch.");

    public bool IsWordProcessAlive() => true;

    public void Dispose()
    {
    }
}
