namespace DocumentGenerator
{
    /// <summary>
    /// Common fluent API shared by all document generators (DOCX, PDF).
    /// Use <see cref="DocumentFactory.Create"/> to obtain an instance based on file extension.
    /// </summary>
    public interface IDocument
    {
        // ── Page setup ────────────────────────────────────────────────

        IDocument SetPageMargins(int topPt = 72, int bottomPt = 72, int leftPt = 72, int rightPt = 72,
            int headerPt = 36, int footerPt = 36);

        IDocument SetHeader(string text);

        IDocument SetFooter(string text, bool includePageNumbers = false);

        IDocument AddPageNumbers();

        // ── Content ───────────────────────────────────────────────────

        IDocument AddParagraph(string text, bool bold = false, bool italic = false,
            int fontSize = 0, string fontName = null, string alignment = null,
            int spacingBeforePt = 0, int spacingAfterPt = 0,
            int indentLeftPt = 0, int indentRightPt = 0);

        IDocument AddHeading(string text, int level = 1, string bookmarkName = null);

        IDocument AddHorizontalRule();

        IDocument AddPageBreak();

        IDocument AddLineBreak();

        IDocument AddSectionSeparator(int spacingPt = 24);

        IDocument AutoSpacing(int spacingAfterPt = 8, int spacingBeforePt = 0,
            double lineSpacingMultiple = 1.15);

        IDocument SetHeadingsStartOnNewPage(int maxLevel = 1);

        // ── Save ──────────────────────────────────────────────────────

        /// <summary>Writes the document to <paramref name="filePath"/>.</summary>
        void Save(string filePath);
    }
}
