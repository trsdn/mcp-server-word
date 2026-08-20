namespace WordMcp.ComInterop;

/// <summary>
/// Constants for Microsoft Word COM interop operations.
/// </summary>
public static class ComInteropConstants
{
    #region Timeouts

    /// <summary>
    /// Timeout for the Word.Quit() operation (30 seconds).
    /// With DisplayAlerts disabled Word quits quickly; this timeout catches hung scenarios.
    /// </summary>
    public static readonly TimeSpan WordQuitTimeout = TimeSpan.FromSeconds(30);

    /// <summary>
    /// Timeout for joining the STA thread after quit.
    /// Must be greater than <see cref="WordQuitTimeout"/> so Dispose waits for shutdown to finish.
    /// </summary>
    public static readonly TimeSpan StaThreadJoinTimeout = WordQuitTimeout + TimeSpan.FromSeconds(15);

    /// <summary>
    /// Default timeout for a single Word operation (5 minutes).
    /// </summary>
    public static readonly TimeSpan DefaultOperationTimeout = TimeSpan.FromMinutes(5);

    /// <summary>
    /// Timeout for starting Word and opening the initial documents.
    /// </summary>
    public static readonly TimeSpan StartupTimeout = TimeSpan.FromMinutes(2);

    /// <summary>
    /// Poll interval used while waiting for a freshly started Word to accept COM calls.
    /// </summary>
    public static readonly TimeSpan WarmupPollInterval = TimeSpan.FromMilliseconds(250);

    #endregion

    #region Word file formats (WdSaveFormat)

    /// <summary>Word 97-2003 binary document (.doc). WdSaveFormat.wdFormatDocument97 = 0.</summary>
    public const int WdFormatDocument97 = 0;

    /// <summary>Plain text (.txt). WdSaveFormat.wdFormatText = 2.</summary>
    public const int WdFormatText = 2;

    /// <summary>Rich Text Format (.rtf). WdSaveFormat.wdFormatRTF = 6.</summary>
    public const int WdFormatRtf = 6;

    /// <summary>Filtered HTML (.html). WdSaveFormat.wdFormatFilteredHTML = 10.</summary>
    public const int WdFormatFilteredHtml = 10;

    /// <summary>Open XML document (.docx). WdSaveFormat.wdFormatXMLDocument = 12.</summary>
    public const int WdFormatXmlDocument = 12;

    /// <summary>Macro-enabled Open XML document (.docm). WdSaveFormat.wdFormatXMLDocumentMacroEnabled = 13.</summary>
    public const int WdFormatXmlDocumentMacroEnabled = 13;

    /// <summary>PDF (.pdf). WdSaveFormat.wdFormatPDF = 17.</summary>
    public const int WdFormatPdf = 17;

    #endregion

    #region Word enumerations

    /// <summary>WdAlertLevel.wdAlertsNone = 0 — suppress all Word dialogs.</summary>
    public const int WdAlertsNone = 0;

    /// <summary>WdStatistic.wdStatisticWords = 0.</summary>
    public const int WdStatisticWords = 0;

    /// <summary>WdStatistic.wdStatisticPages = 2.</summary>
    public const int WdStatisticPages = 2;

    /// <summary>WdStatistic.wdStatisticCharacters = 3.</summary>
    public const int WdStatisticCharacters = 3;

    /// <summary>WdStatistic.wdStatisticParagraphs = 4.</summary>
    public const int WdStatisticParagraphs = 4;

    /// <summary>WdCollapseDirection.wdCollapseEnd = 0.</summary>
    public const int WdCollapseEnd = 0;

    /// <summary>WdCollapseDirection.wdCollapseStart = 1.</summary>
    public const int WdCollapseStart = 1;

    /// <summary>WdParagraphAlignment.wdAlignParagraphLeft = 0.</summary>
    public const int WdAlignParagraphLeft = 0;

    /// <summary>WdParagraphAlignment.wdAlignParagraphCenter = 1.</summary>
    public const int WdAlignParagraphCenter = 1;

    /// <summary>WdParagraphAlignment.wdAlignParagraphRight = 2.</summary>
    public const int WdAlignParagraphRight = 2;

    /// <summary>WdParagraphAlignment.wdAlignParagraphJustify = 3.</summary>
    public const int WdAlignParagraphJustify = 3;

    /// <summary>WdReplace.wdReplaceNone = 0.</summary>
    public const int WdReplaceNone = 0;

    /// <summary>WdReplace.wdReplaceOne = 1.</summary>
    public const int WdReplaceOne = 1;

    /// <summary>WdReplace.wdReplaceAll = 2.</summary>
    public const int WdReplaceAll = 2;

    /// <summary>WdFindWrap.wdFindStop = 0.</summary>
    public const int WdFindStop = 0;

    /// <summary>WdFindWrap.wdFindContinue = 1.</summary>
    public const int WdFindContinue = 1;

    /// <summary>WdFieldType.wdFieldTOC = 13.</summary>
    public const int WdFieldToc = 13;

    /// <summary>WdFieldType.wdFieldNumPages = 26.</summary>
    public const int WdFieldNumPages = 26;

    /// <summary>WdFieldType.wdFieldPage = 33.</summary>
    public const int WdFieldPage = 33;

    /// <summary>WdHeaderFooterIndex.wdHeaderFooterPrimary = 1.</summary>
    public const int WdHeaderFooterPrimary = 1;

    /// <summary>WdInlineShapeType.wdInlineShapePicture = 3.</summary>
    public const int WdInlineShapePicture = 3;

    /// <summary>WdInlineShapeType.wdInlineShapeLinkedPicture = 4.</summary>
    public const int WdInlineShapeLinkedPicture = 4;

    /// <summary>MsoTriState.msoTrue = -1.</summary>
    public const int MsoTrue = -1;

    /// <summary>MsoTriState.msoFalse = 0.</summary>
    public const int MsoFalse = 0;

    /// <summary>WdCaptionLabelID.wdCaptionFigure = -1.</summary>
    public const int WdCaptionFigure = -1;

    /// <summary>WdCaptionPosition.wdCaptionPositionBelow = 1.</summary>
    public const int WdCaptionPositionBelow = 1;

    /// <summary>WdHeaderFooterIndex.wdHeaderFooterFirstPage = 2.</summary>
    public const int WdHeaderFooterFirstPage = 2;

    /// <summary>WdHeaderFooterIndex.wdHeaderFooterEvenPages = 3.</summary>
    public const int WdHeaderFooterEvenPages = 3;

    /// <summary>WdSectionStart.wdSectionContinuous = 0.</summary>
    public const int WdSectionContinuous = 0;

    /// <summary>WdSectionStart.wdSectionNewColumn = 1.</summary>
    public const int WdSectionNewColumn = 1;

    /// <summary>WdSectionStart.wdSectionNewPage = 2.</summary>
    public const int WdSectionNewPage = 2;

    /// <summary>WdSectionStart.wdSectionEvenPage = 3.</summary>
    public const int WdSectionEvenPage = 3;

    /// <summary>WdSectionStart.wdSectionOddPage = 4.</summary>
    public const int WdSectionOddPage = 4;

    /// <summary>WdBreakType.wdSectionBreakNextPage = 2.</summary>
    public const int WdSectionBreakNextPage = 2;

    /// <summary>WdBreakType.wdSectionBreakContinuous = 3.</summary>
    public const int WdSectionBreakContinuous = 3;

    /// <summary>WdBreakType.wdSectionBreakEvenPage = 4.</summary>
    public const int WdSectionBreakEvenPage = 4;

    /// <summary>WdBreakType.wdSectionBreakOddPage = 5.</summary>
    public const int WdSectionBreakOddPage = 5;

    /// <summary>WdOrientation.wdOrientPortrait = 0.</summary>
    public const int WdOrientPortrait = 0;

    /// <summary>WdOrientation.wdOrientLandscape = 1.</summary>
    public const int WdOrientLandscape = 1;

    /// <summary>WdPaperSize.wdPaperA4 = 7.</summary>
    public const int WdPaperA4 = 7;

    /// <summary>WdPaperSize.wdPaperA3 = 6.</summary>
    public const int WdPaperA3 = 6;

    /// <summary>WdPaperSize.wdPaperA5 = 8.</summary>
    public const int WdPaperA5 = 8;

    /// <summary>WdPaperSize.wdPaperLetter = 2.</summary>
    public const int WdPaperLetter = 2;

    /// <summary>WdPaperSize.wdPaperLegal = 5.</summary>
    public const int WdPaperLegal = 5;

    /// <summary>WdPaperSize.wdPaperTabloid = 16.</summary>
    public const int WdPaperTabloid = 16;

    #endregion

    #region Supported extensions

    /// <summary>
    /// Image formats that can be inserted through <c>InlineShapes.AddPicture</c>.
    /// </summary>
    public static readonly string[] SupportedImageExtensions =
        [".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff", ".emf", ".wmf", ".svg"];

    /// <summary>
    /// File extensions this server can open through Word COM automation.
    /// </summary>
    public static readonly string[] SupportedExtensions = [".docx", ".docm", ".doc", ".dotx", ".dotm", ".rtf"];

    /// <summary>
    /// Extensions treated as macro-enabled documents.
    /// </summary>
    public static readonly string[] MacroEnabledExtensions = [".docm", ".dotm"];

    #endregion
}
