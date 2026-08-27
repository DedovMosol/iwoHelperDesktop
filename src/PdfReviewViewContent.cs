namespace ExcelMerger
{
    /// <summary>
    /// Immutable display request for one Review canvas. Unified mode keeps the later page as
    /// its single raster; deletion geometry is a marker layer, never a second line of glyphs
    /// painted over reflowed content.
    /// </summary>
    internal sealed class PdfReviewViewContent
    {
        internal readonly PdfPageRef BasePage;
        internal readonly PdfReviewPage BaseReviewPage;
        internal readonly PdfReviewHighlight BaseHighlight;
        internal readonly PdfPageRef OverlayPage;
        internal readonly PdfReviewHighlight OverlayHighlight;
        internal readonly long Revision;
        internal readonly string Caption;

        internal bool IsComposite { get { return OverlayHighlight != null; } }

        internal PdfReviewViewContent(PdfPageRef basePage, PdfReviewPage baseReviewPage,
            PdfReviewHighlight baseHighlight, PdfPageRef overlayPage,
            PdfReviewHighlight overlayHighlight, long revision, string caption)
        {
            BasePage = basePage == null ? null : basePage.Clone();
            BaseReviewPage = baseReviewPage;
            BaseHighlight = baseHighlight;
            OverlayPage = overlayPage == null ? null : overlayPage.Clone();
            OverlayHighlight = overlayHighlight;
            Revision = revision;
            Caption = caption ?? "";
        }
    }
}
