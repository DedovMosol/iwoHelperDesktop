namespace ExcelMerger
{
    /// <summary>
    /// Immutable display request for one Review canvas. Unified mode uses the later page as
    /// base and projects only deleted fragments from the earlier overlay page.
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

        internal bool IsComposite { get { return OverlayPage != null && OverlayHighlight != null; } }

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
