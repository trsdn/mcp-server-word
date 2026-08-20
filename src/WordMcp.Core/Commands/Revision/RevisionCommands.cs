using System.Runtime.InteropServices;
using WordMcp.ComInterop.Session;
using WordMcp.Core.Models;
using WordMcp.Core.Utilities;

namespace WordMcp.Core.Commands.Revision;

/// <summary>
/// Word COM implementation of <see cref="IRevisionCommands"/>.
/// </summary>
public sealed class RevisionCommands : IRevisionCommands
{
    /// <inheritdoc />
    public RevisionListResult List(IWordBatch batch, string? author = null)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            dynamic doc = ctx.Document;
            dynamic revisions = doc.Revisions;
            int total = (int)revisions.Count;

            var list = new List<RevisionInfo>();

            for (int i = 1; i <= total; i++)
            {
                ct.ThrowIfCancellationRequested();

                var info = Describe(revisions[i], i);

                if (author is not null
                    && !string.Equals(info.Author, author, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                list.Add(info);
            }

            return new RevisionListResult
            {
                TotalCount = total,
                TrackingEnabled = (bool)doc.TrackRevisions,
                Revisions = list
            };
        });
    }

    /// <inheritdoc />
    public RevisionResult Accept(IWordBatch batch, int? index = null)
        => Handle(batch, index, accept: true);

    /// <inheritdoc />
    public RevisionResult Reject(IWordBatch batch, int? index = null)
        => Handle(batch, index, accept: false);

    /// <inheritdoc />
    public RevisionResult SetTracking(IWordBatch batch, bool enabled)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            dynamic doc = ctx.Document;
            doc.TrackRevisions = enabled;

            return new RevisionResult
            {
                HandledCount = 0,
                TotalCount = (int)doc.Revisions.Count,
                TrackingEnabled = (bool)doc.TrackRevisions,
                Message = enabled ? "Change tracking turned on." : "Change tracking turned off."
            };
        });
    }

    private static RevisionResult Handle(IWordBatch batch, int? index, bool accept)
    {
        ArgumentNullException.ThrowIfNull(batch);

        if (index is < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(index), index, "index must be 1 or greater.");
        }

        return batch.Execute((ctx, ct) =>
        {
            dynamic doc = ctx.Document;
            string verb = accept ? "accepted" : "rejected";
            int handled;

            if (index.HasValue)
            {
                int total = (int)doc.Revisions.Count;
                if (index.Value > total)
                {
                    throw new ArgumentOutOfRangeException(
                        nameof(index),
                        $"Revision {index.Value} does not exist. The document has {total} tracked change(s).");
                }

                dynamic revision = doc.Revisions[index.Value];
                if (accept)
                    revision.Accept();
                else
                    revision.Reject();

                handled = 1;
            }
            else
            {
                handled = (int)doc.Revisions.Count;

                if (accept)
                    doc.AcceptAllRevisions();
                else
                    doc.RejectAllRevisions();

                // Document.Revisions covers the body only, the same blind spot field(update-all)
                // works around.
                foreach (dynamic section in doc.Sections)
                {
                    ct.ThrowIfCancellationRequested();
                    handled += HandleHeaderFooter(section.Headers, accept);
                    handled += HandleHeaderFooter(section.Footers, accept);
                }
            }

            return new RevisionResult
            {
                HandledCount = handled,
                TotalCount = (int)doc.Revisions.Count,
                TrackingEnabled = (bool)doc.TrackRevisions,
                Message = index.HasValue
                    ? $"Revision {index.Value} {verb}."
                    : $"{handled} tracked change(s) {verb}."
            };
        });
    }

    private static int HandleHeaderFooter(dynamic headersOrFooters, bool accept)
    {
        int handled = 0;

        foreach (dynamic entry in headersOrFooters)
        {
            dynamic revisions = entry.Range.Revisions;
            int count = (int)revisions.Count;

            if (count == 0)
                continue;

            handled += count;

            if (accept)
                revisions.AcceptAll();
            else
                revisions.RejectAll();
        }

        return handled;
    }

    private static RevisionInfo Describe(dynamic revision, int index)
    {
        var info = new RevisionInfo
        {
            Index = index,
            Type = WordConversions.FromWdRevisionType((int)revision.Type),
            Author = (string?)revision.Author ?? string.Empty,
            Date = ReadDate(revision)
        };

        try
        {
            info.Text = WordConversions.CleanRangeText((string?)revision.Range.Text);
        }
        catch (COMException)
        {
            // Formatting revisions have no text of their own.
        }

        return info;
    }

    private static DateTime? ReadDate(dynamic revision)
    {
        try
        {
            return (DateTime)revision.Date;
        }
        catch (COMException)
        {
            return null;
        }
    }
}
