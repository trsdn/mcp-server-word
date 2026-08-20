using System.Runtime.InteropServices;
using Word = Microsoft.Office.Interop.Word;

namespace WordMcp.ComInterop;

/// <summary>
/// Low-level COM interop helpers for Word automation.
/// </summary>
public static class ComUtilities
{
    /// <summary>
    /// Safely releases a COM object and sets the reference to <c>null</c>.
    /// </summary>
    /// <typeparam name="T">COM wrapper type.</typeparam>
    /// <param name="comObject">The COM object to release.</param>
    /// <remarks>
    /// Release intermediate COM objects (ranges, paragraphs, tables) when iterating collections,
    /// otherwise the WINWORD.EXE process can stay alive after the session ends.
    /// </remarks>
    public static void Release<T>(ref T? comObject) where T : class
    {
        if (comObject != null)
        {
            try
            {
                Marshal.ReleaseComObject(comObject);
            }
            catch (ArgumentException)
            {
                // Not a COM object anymore — already released.
            }
            catch (InvalidComObjectException)
            {
                // RCW already separated from its COM object.
            }

            comObject = null;
        }
    }

    /// <summary>
    /// Fire-and-forget quit of a Word application COM object. Errors are swallowed.
    /// </summary>
    /// <param name="word">The Word.Application COM object, may be <c>null</c>.</param>
    public static void TryQuitWord(Word.Application? word)
    {
        if (word == null) return;

        try
        {
            // SaveChanges: wdDoNotSaveChanges (0) — explicit saves happen through the batch.
            ((dynamic)word).Quit(0);
        }
        catch (COMException)
        {
            // Word already gone or RPC disconnected.
        }
        catch (InvalidComObjectException)
        {
            // RCW already released.
        }
    }
}
