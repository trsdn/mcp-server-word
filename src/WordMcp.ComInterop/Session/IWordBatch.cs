using Microsoft.Extensions.Logging;

namespace WordMcp.ComInterop.Session;

/// <summary>
/// A batch of Word operations that share a single Word instance on a dedicated STA thread.
/// </summary>
/// <remarks>
/// <para><b>Lifecycle:</b> created via <see cref="WordSession.BeginBatch(string[])"/> or
/// <see cref="WordSession.CreateNew(string, bool, TimeSpan?, ILogger?)"/>, used through
/// <c>Execute</c>, persisted via <see cref="Save"/>, and released via <see cref="IDisposable.Dispose"/>.</para>
/// <para><b>Threading:</b> operations are queued and executed serially on one STA thread.
/// This is a COM requirement, not an implementation choice.</para>
/// <example>
/// <code>
/// using var batch = WordSession.BeginBatch("report.docx");
/// var paragraphCount = batch.Execute((ctx, ct) => ctx.Document.Paragraphs.Count);
/// batch.Save();
/// </code>
/// </example>
/// </remarks>
public interface IWordBatch : IDisposable
{
    /// <summary>
    /// Gets the full path of the primary document of this batch.
    /// </summary>
    string DocumentPath { get; }

    /// <summary>
    /// Gets the logger used for diagnostic output.
    /// </summary>
    ILogger Logger { get; }

    /// <summary>
    /// Gets the timeout applied to every <c>Execute</c> call.
    /// </summary>
    TimeSpan OperationTimeout { get; }

    /// <summary>
    /// Gets the WINWORD.EXE process id, when it could be determined.
    /// </summary>
    int? WordProcessId { get; }

    /// <summary>
    /// Executes a COM operation that returns a value on the batch's STA thread.
    /// </summary>
    /// <typeparam name="T">Result type of the operation.</typeparam>
    /// <param name="operation">The operation to execute.</param>
    /// <param name="cancellationToken">Token that cancels waiting for the result.</param>
    /// <returns>The value produced by the operation.</returns>
    /// <exception cref="ObjectDisposedException">The batch has been disposed.</exception>
    /// <exception cref="TimeoutException">The operation exceeded <see cref="OperationTimeout"/>.</exception>
    T Execute<T>(
        Func<WordContext, CancellationToken, T> operation,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Executes a void COM operation on the batch's STA thread.
    /// </summary>
    /// <param name="operation">The operation to execute.</param>
    /// <param name="cancellationToken">Token that cancels waiting for completion.</param>
    /// <exception cref="ObjectDisposedException">The batch has been disposed.</exception>
    /// <exception cref="TimeoutException">The operation exceeded <see cref="OperationTimeout"/>.</exception>
    void Execute(
        Action<WordContext, CancellationToken> operation,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Saves the document. Changes are <b>not</b> saved automatically on dispose.
    /// </summary>
    /// <param name="cancellationToken">Token that cancels waiting for the save.</param>
    void Save(CancellationToken cancellationToken = default);

    /// <summary>
    /// Checks whether the underlying Word process is still alive.
    /// </summary>
    /// <returns><c>true</c> when the process exists and has not exited.</returns>
    bool IsWordProcessAlive();
}
