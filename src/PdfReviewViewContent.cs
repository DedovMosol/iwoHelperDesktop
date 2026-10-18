namespace ExcelMerger
{
    /// <summary>
    /// Immutable display request for one Review canvas. Unified mode keeps the later page as
    /// its single raster; deletions are a marker layer anchored in the later page's own
    /// coordinates, never earlier-page geometry painted over reflowed content.
    /// </summary>
    internal sealed class PdfReviewViewContent
    {
        internal readonly PdfPageRef BasePage;
        internal readonly PdfReviewPage BaseReviewPage;
        internal readonly PdfReviewHighlight BaseHighlight;
        internal readonly PdfReviewHighlight OverlayHighlight;
        internal readonly long Revision;
        internal readonly string Caption;

        internal bool IsComposite { get { return OverlayHighlight != null; } }

        internal PdfReviewViewContent(PdfPageRef basePage, PdfReviewPage baseReviewPage,
            PdfReviewHighlight baseHighlight, PdfReviewHighlight overlayHighlight,
            long revision, string caption)
        {
            BasePage = basePage == null ? null : basePage.Clone();
            BaseReviewPage = baseReviewPage;
            BaseHighlight = baseHighlight;
            OverlayHighlight = overlayHighlight;
            Revision = revision;
            Caption = caption ?? "";
        }
    }
}
