using WordMcp.ComInterop.Session;
using WordMcp.Core.Attributes;
using WordMcp.Core.Models;

namespace WordMcp.Core.Commands.Revision;

/// <summary>
/// Revision operations: listing, accepting and rejecting tracked changes.
/// </summary>
[ServiceCategory("revision", "Revision")]
[McpTool("revision",
    Title = "Revision Operations",
    Description = "Tracked change operations on an open document. "
        + "revision(list, session_id) returns every tracked change with its type, author, date and "
        + "affected text, plus whether tracking is currently on. "
        + "revision(accept, session_id) accepts every change; pass index=1 to accept a single one. "
        + "revision(reject, session_id, index=2) discards a change the same way. "
        + "revision(set-tracking, session_id, enabled=true) turns change tracking on or off, which "
        + "is what makes later edits show up as revisions in the first place. "
        + "Revision indexes are 1-based and shift after every accept or reject, so when handling "
        + "several changes individually work from the highest index downwards, or just accept or "
        + "reject them all at once. "
        + "Accepting or rejecting everything also covers headers and footers, which Word's own "
        + "AcceptAllRevisions does not.")]
public interface IRevisionCommands
{
    /// <summary>
    /// Lists the tracked changes of the document.
    /// </summary>
    /// <param name="batch">The Word batch to operate on.</param>
    /// <param name="author">Restrict the result to one author. All authors when omitted.</param>
    /// <returns>The matching changes.</returns>
    [ServiceAction("list")]
    RevisionListResult List(IWordBatch batch, string? author = null);

    /// <summary>
    /// Accepts one tracked change, or all of them.
    /// </summary>
    /// <param name="batch">The Word batch to operate on.</param>
    /// <param name="index">1-based index of the change. All changes when omitted.</param>
    /// <returns>The number of changes accepted.</returns>
    [ServiceAction("accept")]
    RevisionResult Accept(IWordBatch batch, int? index = null);

    /// <summary>
    /// Rejects one tracked change, or all of them.
    /// </summary>
    /// <param name="batch">The Word batch to operate on.</param>
    /// <param name="index">1-based index of the change. All changes when omitted.</param>
    /// <returns>The number of changes rejected.</returns>
    [ServiceAction("reject")]
    RevisionResult Reject(IWordBatch batch, int? index = null);

    /// <summary>
    /// Turns change tracking on or off.
    /// </summary>
    /// <param name="batch">The Word batch to operate on.</param>
    /// <param name="enabled">Whether edits are recorded as tracked changes.</param>
    /// <returns>The tracking state after the change.</returns>
    [ServiceAction("set-tracking")]
    RevisionResult SetTracking(IWordBatch batch, bool enabled);
}
