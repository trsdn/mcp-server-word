using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading.Channels;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;
using Word = Microsoft.Office.Interop.Word;

namespace WordMcp.ComInterop.Session;

/// <summary>
/// Implementation of <see cref="IWordBatch"/> that owns one Word instance on a dedicated STA thread.
/// </summary>
/// <remarks>
/// <para><b>Word COM threading model:</b></para>
/// <list type="bullet">
/// <item>Each batch runs on exactly one STA thread with a registered OLE message filter.</item>
/// <item>Operations are queued through a channel and executed serially, never in parallel.</item>
/// <item>For parallel work, create separate batches for different documents.</item>
/// </list>
/// </remarks>
internal sealed class WordBatch : IWordBatch
{
    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

    private readonly string _documentPath;
    private readonly string[] _allDocumentPaths;
    private readonly bool _visible;
    private readonly bool _createNewFile;
    private readonly bool _isMacroEnabled;
    private readonly TimeSpan _operationTimeout;
    private readonly ILogger _logger;
    private readonly Channel<Action> _workQueue;
    private readonly Thread _staThread;
    private readonly CancellationTokenSource _shutdownCts;

    private int _disposed;
    private int? _wordProcessId;

    // COM state — touched on the STA thread only.
    private Word.Application? _word;
    private Dictionary<string, Word.Document>? _documents;
    private WordContext? _context;

    /// <summary>
    /// Creates a batch that opens existing documents.
    /// </summary>
    internal WordBatch(
        string[] documentPaths,
        ILogger? logger = null,
        bool visible = false,
        TimeSpan? operationTimeout = null)
        : this(documentPaths, logger, visible, createNewFile: false, isMacroEnabled: false, operationTimeout)
    {
    }

    /// <summary>
    /// Creates a batch that creates a new document and keeps it open.
    /// </summary>
    internal static WordBatch CreateNewDocument(
        string filePath,
        bool isMacroEnabled,
        ILogger? logger = null,
        bool visible = false,
        TimeSpan? operationTimeout = null)
        => new([filePath], logger, visible, createNewFile: true, isMacroEnabled, operationTimeout);

    private WordBatch(
        string[] documentPaths,
        ILogger? logger,
        bool visible,
        bool createNewFile,
        bool isMacroEnabled,
        TimeSpan? operationTimeout)
    {
        if (documentPaths == null || documentPaths.Length == 0)
            throw new ArgumentException("At least one document path is required", nameof(documentPaths));

        _allDocumentPaths = documentPaths;
        _documentPath = documentPaths[0];
        _visible = visible;
        _createNewFile = createNewFile;
        _isMacroEnabled = isMacroEnabled;
        _operationTimeout = operationTimeout ?? ComInteropConstants.DefaultOperationTimeout;
        _logger = logger ?? NullLogger.Instance;
        _shutdownCts = new CancellationTokenSource();

        _workQueue = Channel.CreateUnbounded<Action>(new UnboundedChannelOptions
        {
            SingleReader = true,
            SingleWriter = false
        });

        var started = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);

        _staThread = new Thread(() => RunStaThread(started))
        {
            IsBackground = true,
            Name = $"WordBatch-{Path.GetFileName(_documentPath)}"
        };
        _staThread.SetApartmentState(ApartmentState.STA);
        _staThread.Start();

        try
        {
            if (!started.Task.Wait(ComInteropConstants.StartupTimeout))
            {
                throw new TimeoutException(
                    $"Word did not start within {ComInteropConstants.StartupTimeout.TotalSeconds:N0}s for " +
                    $"'{Path.GetFileName(_documentPath)}'.");
            }
        }
        catch (AggregateException ex) when (ex.InnerException != null)
        {
            Dispose();
            throw ex.InnerException;
        }
        catch
        {
            Dispose();
            throw;
        }
    }

    /// <inheritdoc />
    public string DocumentPath => _documentPath;

    /// <inheritdoc />
    public ILogger Logger => _logger;

    /// <inheritdoc />
    public TimeSpan OperationTimeout => _operationTimeout;

    /// <inheritdoc />
    public int? WordProcessId => _wordProcessId;

    /// <inheritdoc />
    public T Execute<T>(Func<WordContext, CancellationToken, T> operation, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(operation);
        ObjectDisposedException.ThrowIf(Volatile.Read(ref _disposed) == 1, this);

        var completion = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);

        var queued = _workQueue.Writer.TryWrite(() =>
        {
            try
            {
                completion.TrySetResult(operation(_context!, cancellationToken));
            }
#pragma warning disable CA1031 // Marshal every failure back to the calling thread
            catch (Exception ex)
            {
                completion.TrySetException(ex);
            }
#pragma warning restore CA1031
        });

        if (!queued)
        {
            throw new ObjectDisposedException(nameof(WordBatch), "The Word batch is shutting down.");
        }

        return WaitForCompletion(completion.Task, cancellationToken);
    }

    /// <inheritdoc />
    public void Execute(Action<WordContext, CancellationToken> operation, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(operation);

        _ = Execute<object?>((ctx, ct) =>
        {
            operation(ctx, ct);
            return null;
        }, cancellationToken);
    }

    /// <inheritdoc />
    public void Save(CancellationToken cancellationToken = default)
    {
        Execute((ctx, ct) =>
        {
            foreach (var doc in _documents!.Values)
            {
                ((dynamic)doc).Save();
            }
        }, cancellationToken);
    }

    /// <inheritdoc />
    public bool IsWordProcessAlive()
    {
        if (!_wordProcessId.HasValue)
            return false;

        try
        {
            using var process = Process.GetProcessById(_wordProcessId.Value);
            return !process.HasExited;
        }
        catch (ArgumentException)
        {
            return false; // Process no longer exists
        }
        catch (InvalidOperationException)
        {
            return false;
        }
    }

    /// <inheritdoc />
    public void Dispose()
    {
        if (Interlocked.CompareExchange(ref _disposed, 1, 0) != 0)
            return;

        _workQueue.Writer.TryComplete();

        if (!_staThread.Join(ComInteropConstants.StaThreadJoinTimeout))
        {
            _logger.LogWarning(
                "Word STA thread did not shut down within {Timeout}s for '{File}'. Forcing cleanup.",
                ComInteropConstants.StaThreadJoinTimeout.TotalSeconds,
                Path.GetFileName(_documentPath));

            ForceKillWord();
        }

        _shutdownCts.Cancel();
        _shutdownCts.Dispose();
    }

    private T WaitForCompletion<T>(Task<T> task, CancellationToken cancellationToken)
    {
        bool completed;

        try
        {
            completed = task.Wait(_operationTimeout, cancellationToken);
        }
        catch (AggregateException)
        {
            // Wait wraps failures; GetResult below rethrows the original exception so callers see
            // the actual COM or argument error instead of an AggregateException.
            completed = true;
        }

        if (!completed)
        {
            throw new TimeoutException(
                $"Word operation on '{Path.GetFileName(_documentPath)}' exceeded " +
                $"{_operationTimeout.TotalSeconds:N0}s. Word may be showing a dialog or is unresponsive.");
        }

        return task.GetAwaiter().GetResult();
    }

    private void RunStaThread(TaskCompletionSource started)
    {
        try
        {
            OleMessageFilter.Register();
            StartWordAndOpenDocuments();
            started.SetResult();
        }
#pragma warning disable CA1031 // Startup failures must surface on the calling thread
        catch (Exception ex)
        {
            started.TrySetException(ex);
            CleanupCom();
            OleMessageFilter.Revoke();
            return;
        }
#pragma warning restore CA1031

        try
        {
            var reader = _workQueue.Reader;
            while (true)
            {
                if (!reader.WaitToReadAsync().AsTask().GetAwaiter().GetResult())
                    break; // Channel completed — batch is being disposed

                while (reader.TryRead(out var workItem))
                {
                    workItem();
                }
            }
        }
        finally
        {
            CleanupCom();
            OleMessageFilter.Revoke();
        }
    }

    private void StartWordAndOpenDocuments()
    {
        Type? wordType = Type.GetTypeFromProgID("Word.Application");
        if (wordType == null)
        {
            throw new InvalidOperationException(
                "Microsoft Word is not installed or not registered on this system.");
        }

#pragma warning disable IL2072 // COM activation via ProgID is inherently dynamic
        var app = (Word.Application)Activator.CreateInstance(wordType)!;
#pragma warning restore IL2072

        WaitUntilResponsive(app);

        // Unlike PowerPoint, Word allows hiding the application window entirely.
        ((dynamic)app).Visible = _visible;
        ((dynamic)app).DisplayAlerts = ComInteropConstants.WdAlertsNone;
        ((dynamic)app).ScreenUpdating = _visible;

        var documents = new Dictionary<string, Word.Document>(StringComparer.OrdinalIgnoreCase);
        Word.Document? primary = null;

        foreach (var path in _allDocumentPaths)
        {
            string normalizedPath = Path.GetFullPath(path);
            Word.Document doc = _createNewFile
                ? CreateDocument(app, normalizedPath)
                : OpenDocument(app, normalizedPath);

            documents[normalizedPath] = doc;

            if (string.Equals(normalizedPath, Path.GetFullPath(_documentPath), StringComparison.OrdinalIgnoreCase))
            {
                primary = doc;
            }
        }

        _word = app;
        _documents = documents;
        _context = new WordContext(Path.GetFullPath(_documentPath), app, primary!);

        CaptureWordProcessId(app);
    }

    /// <summary>
    /// Blocks until a freshly started Word answers a trivial call without rejecting it.
    /// </summary>
    /// <remarks>
    /// Word rejects COM calls for several seconds after activation. Issuing document work during
    /// that window makes each call wait on the OLE message filter, and the resulting delay lets
    /// Word's AutoSave claim a new document for OneDrive before it can be saved locally — the
    /// document then lives in the cloud and every later local save fails or blocks on a dialog.
    /// Warming Word up first keeps <see cref="CreateDocument"/> fast enough to win that race.
    /// </remarks>
    private static void WaitUntilResponsive(Word.Application app)
    {
        var deadline = DateTime.UtcNow + ComInteropConstants.StartupTimeout;
        Exception? last = null;

        while (DateTime.UtcNow < deadline)
        {
            try
            {
                _ = (int)((dynamic)app).Documents.Count;
                return;
            }
            catch (COMException ex)
            {
                last = ex;
                Thread.Sleep(ComInteropConstants.WarmupPollInterval);
            }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException ex)
            {
                last = ex;
                Thread.Sleep(ComInteropConstants.WarmupPollInterval);
            }
        }

        throw new TimeoutException(
            $"Word did not become responsive within {ComInteropConstants.StartupTimeout.TotalSeconds:N0}s.",
            last);
    }

    private Word.Document CreateDocument(Word.Application app, string normalizedPath)
    {
        string? directory = Path.GetDirectoryName(normalizedPath);
        if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
        {
            throw new DirectoryNotFoundException(
                $"Directory does not exist: '{directory}'. Create it before creating Word documents.");
        }

        EmptyDocumentFactory.Create(normalizedPath, _isMacroEnabled);

        Word.Document doc = OpenDocument(app, normalizedPath);

        return doc;
    }

    private static Word.Document OpenDocument(Word.Application app, string normalizedPath)
    {
        bool isIrm = FileAccessValidator.IsIrmProtected(normalizedPath);

        if (isIrm)
        {
            // IRM-protected files need a visible Word so the user can authenticate.
            ((dynamic)app).Visible = true;
        }
        else
        {
            FileAccessValidator.ValidateFileNotLocked(normalizedPath);
        }

        try
        {
            var doc = isIrm
                ? ((dynamic)app).Documents.Open(normalizedPath, ReadOnly: true, Visible: true)
                : ((dynamic)app).Documents.Open(normalizedPath, ReadOnly: false, AddToRecentFiles: false);

            return (Word.Document)doc;
        }
        catch (COMException ex)
        {
            throw FileAccessValidator.CreateFileLockedError(normalizedPath, ex);
        }
    }

    private void CaptureWordProcessId(Word.Application app)
    {
        try
        {
            int hwnd = ((dynamic)app).ActiveWindow.Hwnd;
            if (hwnd != 0 && GetWindowThreadProcessId(new IntPtr(hwnd), out uint processId) != 0 && processId != 0)
            {
                _wordProcessId = (int)processId;
                _logger.LogDebug("Captured Word process id {ProcessId}", _wordProcessId);
                return;
            }
        }
        catch (COMException ex)
        {
            _logger.LogDebug(ex, "Could not read Word window handle; force-kill will be unavailable.");
        }
        catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException ex)
        {
            _logger.LogDebug(ex, "Word does not expose ActiveWindow; force-kill will be unavailable.");
        }

        _logger.LogWarning(
            "Could not determine the Word process id. Force-kill is disabled for this session " +
            "to avoid terminating unrelated Word instances.");
    }

    private void CleanupCom()
    {
        if (_documents != null)
        {
            foreach (var doc in _documents.Values)
            {
                try
                {
                    // wdDoNotSaveChanges (0) — unsaved changes are discarded by design.
                    ((dynamic)doc).Close(0);
                }
                catch (COMException)
                {
                    // Document already closed or Word disconnected.
                }

                var comDoc = doc;
                ComUtilities.Release(ref comDoc!);
            }

            _documents.Clear();
            _documents = null;
        }

        ComUtilities.TryQuitWord(_word);
        ComUtilities.Release(ref _word);

        _context = null;
    }

    private void ForceKillWord()
    {
        if (!_wordProcessId.HasValue)
            return;

        try
        {
            using var process = Process.GetProcessById(_wordProcessId.Value);
            if (!process.HasExited)
            {
                process.Kill(entireProcessTree: true);
                _logger.LogWarning("Force-killed unresponsive Word process {ProcessId}", _wordProcessId);
            }
        }
        catch (ArgumentException)
        {
            // Process already gone.
        }
        catch (InvalidOperationException)
        {
            // Process already exited.
        }
        catch (System.ComponentModel.Win32Exception ex)
        {
            _logger.LogWarning(ex, "Failed to force-kill Word process {ProcessId}", _wordProcessId);
        }
    }
}
