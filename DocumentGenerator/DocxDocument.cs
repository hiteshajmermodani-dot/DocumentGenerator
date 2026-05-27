using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;

namespace DocumentGenerator
{
    /// <summary>
    /// A fluent API for generating .docx files without any third-party libraries.
    /// </summary>
    public class DocxDocument : IDocument
    {
        private readonly List<string> _bodyElements = new List<string>();
        private readonly List<DocxImage> _images = new List<DocxImage>();
        private readonly List<string> _relationships = new List<string>();
        private DocxHeaderFooter _header;
        private DocxHeaderFooter _footer;
        private bool _includePageNumbers;
        private int _relationshipId = 1;
        private int _pageMarginTop = 1440;
        private int _pageMarginBottom = 1440;
        private int _pageMarginLeft = 1440;
        private int _pageMarginRight = 1440;
        private int _pageMarginHeader = 720;
        private int _pageMarginFooter = 720;
        private int _defaultSpacingAfterPt;
        private int _defaultSpacingBeforePt;
        private int _lineSpacingTwips;
        private bool _pageBreakBeforeHeadings;
        private int _pageBreakBeforeLevel = 1;

        // Page-number tracking for TOC PAGEREF cache values
        private int _currentPage = 1;
        private bool _justAddedPageBreak = false;
        private readonly Dictionary<string, int> _bookmarkPages = new Dictionary<string, int>();
        private DocxTableOfContents _pendingToc;
        private const string TocPlaceholder = "\x00__TOC_PLACEHOLDER__\x00";

        // Content-height accumulator for estimating page overflows
        // Standard letter page = 12240 twips tall; default margins 1440 top+bottom = 9360 usable
        private int _usedTwips = 0;
        private int _usablePageHeightTwips = 12240 - 1440 - 1440; // recalculated on SetPageMargins

        // Threshold in twips: if remainder after page overflow is less than this,
        // we consider the page "just started" and won't add another page break.
        // ~2 inches = 2880 twips gives enough margin to absorb table row estimation variance.
        private const int NewPageThresholdTwips = 2880;

        /// <summary>Accumulates estimated content height and auto-advances the page counter.</summary>
        private void ConsumeHeight(int twips)
        {
            _usedTwips += twips;
            bool overflowed = false;
            while (_usedTwips >= _usablePageHeightTwips)
            {
                _usedTwips -= _usablePageHeightTwips;
                _currentPage++;
                overflowed = true;
            }
            // If we just overflowed to a new page with very little content remaining,
            // treat it as being at top of page so ForcePageBreak won't double-count
            if (overflowed && _usedTwips < NewPageThresholdTwips)
                _justAddedPageBreak = true;
            else
                _justAddedPageBreak = false;
        }

        /// <summary>
        /// Forces a logical page break only if meaningful content has been placed on the current page.
        /// Avoids double-counting when ConsumeHeight already overflowed into a fresh page.
        /// </summary>
        private void ForcePageBreak()
        {
            if (!_justAddedPageBreak && _usedTwips >= NewPageThresholdTwips)
            {
                _usedTwips = 0;
                _currentPage++;
            }
            else
            {
                _usedTwips = 0;
            }
            _justAddedPageBreak = true;
        }

        private const string WmlNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        private const string RelNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";
        private const string DrawingNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main";
        private const string WpNamespace = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
        private const string PicNamespace = "http://schemas.openxmlformats.org/drawingml/2006/picture";
        private const string RNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        /// <summary>
        /// Creates a new empty DocxDocument.
        /// </summary>
        public DocxDocument()
        {
        }

        /// <summary>
        /// Sets page margins in points. Default is 72pt (1 inch) for all sides.
        /// </summary>
        /// <param name="topPt">Top margin in points.</param>
        /// <param name="bottomPt">Bottom margin in points.</param>
        /// <param name="leftPt">Left margin in points.</param>
        /// <param name="rightPt">Right margin in points.</param>
        /// <param name="headerPt">Header distance from edge in points.</param>
        /// <param name="footerPt">Footer distance from edge in points.</param>
        public DocxDocument SetPageMargins(int topPt = 72, int bottomPt = 72, int leftPt = 72, int rightPt = 72, int headerPt = 36, int footerPt = 36)
        {
            _pageMarginTop = topPt * 20;
            _pageMarginBottom = bottomPt * 20;
            _pageMarginLeft = leftPt * 20;
            _pageMarginRight = rightPt * 20;
            _pageMarginHeader = headerPt * 20;
            _pageMarginFooter = footerPt * 20;
            // Recalculate usable page height: standard letter page = 12240 twips (8.5x11in)
            _usablePageHeightTwips = 12240 - _pageMarginTop - _pageMarginBottom;
            return this;
        }

        /// <summary>
        /// Automatically adjusts spacing between all paragraphs, headings, and elements 
        /// for a clean, evenly-spaced layout. Applies default before/after spacing and line spacing.
        /// </summary>
        /// <param name="spacingAfterPt">Default space after each paragraph in points. Default is 8pt.</param>
        /// <param name="spacingBeforePt">Default space before each paragraph in points. Default is 0pt.</param>
        /// <param name="lineSpacingMultiple">Line spacing multiplier (1.0 = single, 1.15, 1.5, 2.0 = double). Default is 1.15.</param>
        public DocxDocument AutoSpacing(int spacingAfterPt = 8, int spacingBeforePt = 0, double lineSpacingMultiple = 1.15)
        {
            _defaultSpacingAfterPt = spacingAfterPt;
            _defaultSpacingBeforePt = spacingBeforePt;
            _lineSpacingTwips = (int)(lineSpacingMultiple * 240);
            return this;
        }

        /// <summary>
        /// When enabled, each heading at or below the specified level will automatically
        /// start on a new page. This ensures sections don't share pages.
        /// </summary>
        /// <param name="maxLevel">Headings at this level or higher (1=largest) will start on a new page. Default is 1 (only level-1 headings).</param>
        public DocxDocument SetHeadingsStartOnNewPage(int maxLevel = 1)
        {
            _pageBreakBeforeHeadings = true;
            _pageBreakBeforeLevel = maxLevel;
            return this;
        }

        /// <summary>
        /// Adds a paragraph with the specified text.
        /// </summary>
        public DocxDocument AddParagraph(string text, bool bold = false, bool italic = false, int fontSize = 0, string fontName = null, string alignment = null, int spacingBeforePt = 0, int spacingAfterPt = 0, int indentLeftPt = 0, int indentRightPt = 0)
        {
            var runProps = BuildRunProperties(bold, italic, fontSize, fontName);
            var paraProps = BuildParagraphProperties(alignment, spacingBeforePt, spacingAfterPt, indentLeftPt, indentRightPt);

            _bodyElements.Add(
                $"<w:p>{paraProps}<w:r>{runProps}<w:t xml:space=\"preserve\">{EscapeXml(text)}</w:t></w:r></w:p>");

            // Estimate height: account for text wrapping (~80 chars per line at default width/font)
            int linePt = fontSize > 0 ? fontSize / 2 : 12;
            int charsPerLine = fontSize > 0 ? Math.Max(20, 1440 / Math.Max(1, fontSize / 2)) : 80;
            int lineCount = string.IsNullOrEmpty(text) ? 1 : Math.Max(1, (int)Math.Ceiling((double)text.Length / charsPerLine));
            int totalPt = linePt * lineCount + spacingBeforePt + spacingAfterPt + _defaultSpacingAfterPt + _defaultSpacingBeforePt;
            ConsumeHeight(totalPt * 20);
            return this;
        }

        /// <summary>
        /// Adds a heading paragraph (level 1-9).
        /// </summary>
        public DocxDocument AddHeading(string text, int level = 1, string bookmarkName = null)
        {
            if (level < 1) level = 1;
            if (level > 9) level = 9;

            // Force a logical page break if configured, using height-based tracking
            if (_pageBreakBeforeHeadings && level <= _pageBreakBeforeLevel)
                ForcePageBreak();

            // Record page number for this bookmark BEFORE consuming heading height
            if (!string.IsNullOrEmpty(bookmarkName))
                _bookmarkPages[bookmarkName] = _currentPage;

            // Estimate heading height: level 1 = 28pt font, shrinks by 2pt per level
            int headingPt = Math.Max(12, 28 - (level - 1) * 2) + _defaultSpacingAfterPt + _defaultSpacingBeforePt;
            ConsumeHeight(headingPt * 20);

            string bookmarkId = Math.Abs(Guid.NewGuid().GetHashCode()).ToString();
            string anchorName = "_Toc_" + bookmarkId;

            var sb = new StringBuilder();
            sb.Append("<w:p>");
            sb.Append("<w:pPr>");
            sb.Append($"<w:pStyle w:val=\"Heading{level}\"/>");
            if (_pageBreakBeforeHeadings && level <= _pageBreakBeforeLevel)
                sb.Append("<w:pageBreakBefore/>");
            sb.Append("</w:pPr>");

            // TOC bookmark
            sb.Append($"<w:bookmarkStart w:id=\"{bookmarkId}\" w:name=\"{anchorName}\"/>");

            // User-specified bookmark for navigation
            string userBookmarkId = null;
            if (!string.IsNullOrEmpty(bookmarkName))
            {
                userBookmarkId = Math.Abs(Guid.NewGuid().GetHashCode()).ToString();
                sb.Append($"<w:bookmarkStart w:id=\"{userBookmarkId}\" w:name=\"{EscapeXml(bookmarkName)}\"/>");
            }

            sb.Append($"<w:r><w:rPr><w:b/><w:sz w:val=\"{(48 - (level - 1) * 4)}\"/></w:rPr>");
            sb.Append($"<w:t xml:space=\"preserve\">{EscapeXml(text)}</w:t></w:r>");

            sb.Append($"<w:bookmarkEnd w:id=\"{bookmarkId}\"/>");
            if (userBookmarkId != null)
                sb.Append($"<w:bookmarkEnd w:id=\"{userBookmarkId}\"/>");

            sb.Append("</w:p>");

            _bodyElements.Add(sb.ToString());
            return this;
        }

        /// <summary>
        /// Adds a page break.
        /// </summary>
        public DocxDocument AddPageBreak()
        {
            _bodyElements.Add(
                "<w:p><w:r><w:br w:type=\"page\"/></w:r></w:p>");
            // Always force to next page regardless of current position
            _usedTwips = 0;
            _currentPage++;
            _justAddedPageBreak = true;
            return this;
        }

        /// <summary>
        /// Adds a section separator with vertical spacing. Content flows naturally
        /// without forcing a new page, filling available space before moving to the next page.
        /// </summary>
        /// <param name="spacingPt">Spacing in points between sections. Default is 24pt.</param>
        public DocxDocument AddSectionSeparator(int spacingPt = 24)
        {
            _bodyElements.Add(
                $"<w:p><w:pPr><w:spacing w:before=\"{spacingPt * 20}\" w:after=\"{spacingPt * 20}\"/></w:pPr></w:p>");
            ConsumeHeight(spacingPt * 2 * 20);
            return this;
        }

        /// <summary>
        /// Adds a line break within the current flow.
        /// </summary>
        public DocxDocument AddLineBreak()
        {
            _bodyElements.Add(
                "<w:p><w:r><w:br/></w:r></w:p>");
            return this;
        }

        /// <summary>
        /// Adds a horizontal rule (border paragraph).
        /// </summary>
        public DocxDocument AddHorizontalRule()
        {
            _bodyElements.Add(
                "<w:p><w:pPr><w:pBdr><w:bottom w:val=\"single\" w:sz=\"12\" w:space=\"1\" w:color=\"000000\"/></w:pBdr></w:pPr></w:p>");
            ConsumeHeight(240);
            return this;
        }

        /// <summary>
        /// Adds a table using the DocxTable builder.
        /// </summary>
        public DocxDocument AddTable(DocxTable table)
        {
            _bodyElements.Add(table.ToXml());
            // Estimate: each row ~420 twips (21pt) to account for text, cell padding, and borders
            ConsumeHeight(table.RowCount * 420);
            return this;
        }

        /// <summary>
        /// Adds a bitmap image from a file path.
        /// </summary>
        /// <param name="alignment">Paragraph alignment: left, center, right.</param>
        /// <param name="spacingBeforePt">Space before image in points.</param>
        /// <param name="spacingAfterPt">Space after image in points.</param>
        public DocxDocument AddImage(string imagePath, int widthEmu = 5000000, int heightEmu = 3000000, string alignment = null, int spacingBeforePt = 0, int spacingAfterPt = 0)
        {
            var image = new DocxImage(imagePath, widthEmu, heightEmu, alignment, spacingBeforePt, spacingAfterPt);
            string rId = "rId" + (++_relationshipId);
            image.RelationshipId = rId;
            _images.Add(image);

            _relationships.Add(
                $"<Relationship Id=\"{rId}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"media/{image.FileName}\"/>");

            _bodyElements.Add(image.ToXml());
            // Convert EMU to twips: 1 twip = 635 EMU; add spacing
            ConsumeHeight(heightEmu / 635 + (spacingBeforePt + spacingAfterPt) * 20);
            return this;
        }

        /// <summary>
        /// Adds a bitmap image from a byte array.
        /// </summary>
        public DocxDocument AddImage(byte[] imageData, string fileName, int widthEmu = 5000000, int heightEmu = 3000000, string alignment = null, int spacingBeforePt = 0, int spacingAfterPt = 0)
        {
            var image = new DocxImage(imageData, fileName, widthEmu, heightEmu, alignment, spacingBeforePt, spacingAfterPt);
            string rId = "rId" + (++_relationshipId);
            image.RelationshipId = rId;
            _images.Add(image);

            _relationships.Add(
                $"<Relationship Id=\"{rId}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"media/{image.FileName}\"/>");

            _bodyElements.Add(image.ToXml());
            // Convert EMU to twips: 1 twip = 635 EMU; add spacing
            ConsumeHeight(heightEmu / 635 + (spacingBeforePt + spacingAfterPt) * 20);
            return this;
        }

        /// <summary>
        /// Sets the document header.
        /// </summary>
        public DocxDocument SetHeader(DocxHeaderFooter header)
        {
            _header = header;
            return this;
        }

        /// <summary>
        /// Sets the document footer.
        /// </summary>
        public DocxDocument SetFooter(DocxHeaderFooter footer)
        {
            _footer = footer;
            return this;
        }

        /// <summary>
        /// Enables page numbers in the footer.
        /// </summary>
        public DocxDocument AddPageNumbers()
        {
            _includePageNumbers = true;
            return this;
        }

        /// <summary>
        /// Adds a structured Table of Contents at the current position.
        /// The TOC XML is resolved after all headings are added so page numbers are correct.
        /// </summary>
        public DocxDocument AddTableOfContents(DocxTableOfContents toc)
        {
            _pendingToc = toc;
            _bodyElements.Add(TocPlaceholder);
            // TOC is on page 1; headings follow after a page break
            _justAddedPageBreak = false;
            return this;
        }

        /// <summary>
        /// Adds a named bookmark at the current position.
        /// </summary>
        public DocxDocument AddBookmark(string name)
        {
            string id = Math.Abs(Guid.NewGuid().GetHashCode()).ToString();
            _bodyElements.Add(
                $"<w:p><w:bookmarkStart w:id=\"{id}\" w:name=\"{EscapeXml(name)}\"/><w:bookmarkEnd w:id=\"{id}\"/></w:p>");
            return this;
        }

        /// <summary>
        /// Saves the document to the specified file path.
        /// </summary>
        public void Save(string filePath)
        {
            string tempDir = Path.Combine(Path.GetTempPath(), "docx_gen_" + Guid.NewGuid().ToString("N"));

            try
            {
                Directory.CreateDirectory(tempDir);
                Directory.CreateDirectory(Path.Combine(tempDir, "_rels"));
                Directory.CreateDirectory(Path.Combine(tempDir, "word"));
                Directory.CreateDirectory(Path.Combine(tempDir, "word", "_rels"));

                if (_images.Count > 0)
                    Directory.CreateDirectory(Path.Combine(tempDir, "word", "media"));

                WriteContentTypes(tempDir);
                WriteRootRelationships(tempDir);
                WriteDocumentRelationships(tempDir);
                WriteDocument(tempDir);
                WriteImages(tempDir);
                WriteHeaderFooter(tempDir);
                WriteSettings(tempDir);
                WriteStyles(tempDir);

                if (File.Exists(filePath))
                    File.Delete(filePath);

                ZipFile.CreateFromDirectory(tempDir, filePath);
            }
            finally
            {
                if (Directory.Exists(tempDir))
                    Directory.Delete(tempDir, true);
            }
        }

        /// <summary>
        /// Saves the document to a stream.
        /// </summary>
        public void Save(Stream stream)
        {
            string tempPath = Path.Combine(Path.GetTempPath(), "docx_temp_" + Guid.NewGuid().ToString("N") + ".docx");
            try
            {
                Save(tempPath);
                byte[] bytes = File.ReadAllBytes(tempPath);
                stream.Write(bytes, 0, bytes.Length);
            }
            finally
            {
                if (File.Exists(tempPath))
                    File.Delete(tempPath);
            }
        }

        private void WriteContentTypes(string tempDir)
        {
            var sb = new StringBuilder();
            sb.AppendLine("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
            sb.AppendLine("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");
            sb.AppendLine("  <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>");
            sb.AppendLine("  <Default Extension=\"xml\" ContentType=\"application/xml\"/>");
            sb.AppendLine("  <Default Extension=\"png\" ContentType=\"image/png\"/>");
            sb.AppendLine("  <Default Extension=\"jpg\" ContentType=\"image/jpeg\"/>");
            sb.AppendLine("  <Default Extension=\"jpeg\" ContentType=\"image/jpeg\"/>");
            sb.AppendLine("  <Default Extension=\"bmp\" ContentType=\"image/bmp\"/>");
            sb.AppendLine("  <Default Extension=\"gif\" ContentType=\"image/gif\"/>");
            sb.AppendLine("  <Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>");
            sb.AppendLine("  <Override PartName=\"/word/settings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml\"/>");
            sb.AppendLine("  <Override PartName=\"/word/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml\"/>");

            if (_header != null)
                sb.AppendLine("  <Override PartName=\"/word/header1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml\"/>");
            if (_footer != null || _includePageNumbers)
                sb.AppendLine("  <Override PartName=\"/word/footer1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml\"/>");

            sb.AppendLine("</Types>");
            File.WriteAllText(Path.Combine(tempDir, "[Content_Types].xml"), sb.ToString());
        }

        private void WriteRootRelationships(string tempDir)
        {
            File.WriteAllText(Path.Combine(tempDir, "_rels", ".rels"),
@"<?xml version=""1.0"" encoding=""UTF-8""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
  <Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"" Target=""word/document.xml""/>
</Relationships>");
        }

        private void WriteDocumentRelationships(string tempDir)
        {
            var sb = new StringBuilder();
            sb.AppendLine("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
            sb.AppendLine($"<Relationships xmlns=\"{RelNamespace}\">");

            foreach (var rel in _relationships)
                sb.AppendLine("  " + rel);

            if (_header != null)
            {
                string hId = "rId" + (++_relationshipId);
                _header.RelationshipId = hId;
                sb.AppendLine($"  <Relationship Id=\"{hId}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/header\" Target=\"header1.xml\"/>");
            }

            if (_footer != null || _includePageNumbers)
            {
                string fId = "rId" + (++_relationshipId);
                if (_footer != null) _footer.RelationshipId = fId;
                else
                {
                    _footer = new DocxHeaderFooter();
                    _footer.RelationshipId = fId;
                }
                if (_includePageNumbers) _footer.IncludePageNumber = true;
                sb.AppendLine($"  <Relationship Id=\"{fId}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer\" Target=\"footer1.xml\"/>");
            }

            string sId = "rId" + (++_relationshipId);
            sb.AppendLine($"  <Relationship Id=\"{sId}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings\" Target=\"settings.xml\"/>");

            string stId = "rId" + (++_relationshipId);
            sb.AppendLine($"  <Relationship Id=\"{stId}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>");

            sb.AppendLine("</Relationships>");
            File.WriteAllText(Path.Combine(tempDir, "word", "_rels", "document.xml.rels"), sb.ToString());
        }

        private void WriteDocument(string tempDir)
        {
            var sb = new StringBuilder();
            sb.AppendLine("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
            sb.Append("<w:document xmlns:w=\"" + WmlNamespace + "\"");
            sb.Append(" xmlns:r=\"" + RNamespace + "\"");
            sb.Append(" xmlns:wp=\"" + WpNamespace + "\"");
            sb.Append(" xmlns:a=\"" + DrawingNamespace + "\"");
            sb.AppendLine(" xmlns:pic=\"" + PicNamespace + "\">");
            sb.AppendLine("  <w:body>");

            foreach (var element in _bodyElements)
            {
                if (element == TocPlaceholder && _pendingToc != null)
                    sb.AppendLine("    " + _pendingToc.ToXml(_bookmarkPages));
                else
                    sb.AppendLine("    " + element);
            }

            // Section properties for page margins and header/footer references
            sb.AppendLine("    <w:sectPr>");
            if (_header != null)
                sb.AppendLine($"      <w:headerReference w:type=\"default\" r:id=\"{_header.RelationshipId}\"/>");
            if (_footer != null)
                sb.AppendLine($"      <w:footerReference w:type=\"default\" r:id=\"{_footer.RelationshipId}\"/>");
            sb.AppendLine($"      <w:pgMar w:top=\"{_pageMarginTop}\" w:right=\"{_pageMarginRight}\" w:bottom=\"{_pageMarginBottom}\" w:left=\"{_pageMarginLeft}\" w:header=\"{_pageMarginHeader}\" w:footer=\"{_pageMarginFooter}\"/>");
            sb.AppendLine("    </w:sectPr>");

            sb.AppendLine("  </w:body>");
            sb.AppendLine("</w:document>");
            File.WriteAllText(Path.Combine(tempDir, "word", "document.xml"), sb.ToString());
        }

        private void WriteImages(string tempDir)
        {
            foreach (var image in _images)
            {
                string mediaPath = Path.Combine(tempDir, "word", "media", image.FileName);
                File.WriteAllBytes(mediaPath, image.GetImageData());
            }
        }

        private void WriteHeaderFooter(string tempDir)
        {
            if (_header != null)
                File.WriteAllText(Path.Combine(tempDir, "word", "header1.xml"), _header.ToHeaderXml());

            if (_footer != null)
                File.WriteAllText(Path.Combine(tempDir, "word", "footer1.xml"), _footer.ToFooterXml());
        }

        private void WriteSettings(string tempDir)
        {
            var sb = new StringBuilder();
            sb.AppendLine("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
            sb.Append("<w:settings xmlns:w=\"" + WmlNamespace + "\">");
            sb.Append("<w:updateFields w:val=\"true\"/>");
            sb.Append("</w:settings>");
            File.WriteAllText(Path.Combine(tempDir, "word", "settings.xml"), sb.ToString());
        }

        private void WriteStyles(string tempDir)
        {
            var sb = new StringBuilder();
            sb.AppendLine("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
            sb.Append("<w:styles xmlns:w=\"" + WmlNamespace + "\">");

            // Default paragraph spacing (applies to all paragraphs as base style)
            sb.Append("<w:docDefaults>");
            sb.Append("<w:pPrDefault><w:pPr>");
            if (_defaultSpacingAfterPt > 0 || _defaultSpacingBeforePt > 0 || _lineSpacingTwips > 0)
            {
                sb.Append("<w:spacing");
                if (_defaultSpacingBeforePt > 0) sb.Append($" w:before=\"{_defaultSpacingBeforePt * 20}\"");
                if (_defaultSpacingAfterPt > 0) sb.Append($" w:after=\"{_defaultSpacingAfterPt * 20}\"");
                if (_lineSpacingTwips > 0) sb.Append($" w:line=\"{_lineSpacingTwips}\" w:lineRule=\"auto\"");
                sb.Append("/>");
            }
            sb.Append("</w:pPr></w:pPrDefault>");
            sb.Append("</w:docDefaults>");

            // Heading styles
            for (int i = 1; i <= 9; i++)
            {
                sb.Append($"<w:style w:type=\"paragraph\" w:styleId=\"Heading{i}\">");
                sb.Append($"<w:name w:val=\"heading {i}\"/>");
                sb.Append($"<w:pPr><w:outlineLvl w:val=\"{i - 1}\"/></w:pPr>");
                sb.Append("</w:style>");
            }

            // TOC styles
            for (int i = 1; i <= 3; i++)
            {
                sb.Append($"<w:style w:type=\"paragraph\" w:styleId=\"TOC{i}\">");
                sb.Append($"<w:name w:val=\"toc {i}\"/>");
                sb.Append("</w:style>");
            }

            // Hyperlink character style
            sb.Append("<w:style w:type=\"character\" w:styleId=\"Hyperlink\">");
            sb.Append("<w:name w:val=\"Hyperlink\"/>");
            sb.Append("<w:rPr><w:color w:val=\"0563C1\"/><w:u w:val=\"single\"/></w:rPr>");
            sb.Append("</w:style>");

            sb.Append("</w:styles>");
            File.WriteAllText(Path.Combine(tempDir, "word", "styles.xml"), sb.ToString());
        }

        private static string BuildRunProperties(bool bold, bool italic, int fontSize, string fontName)
        {
            if (!bold && !italic && fontSize == 0 && fontName == null)
                return string.Empty;

            var sb = new StringBuilder("<w:rPr>");
            if (bold) sb.Append("<w:b/>");
            if (italic) sb.Append("<w:i/>");
            if (fontSize > 0) sb.Append($"<w:sz w:val=\"{fontSize * 2}\"/>");
            if (fontName != null) sb.Append($"<w:rFonts w:ascii=\"{EscapeXml(fontName)}\" w:hAnsi=\"{EscapeXml(fontName)}\"/>");
            sb.Append("</w:rPr>");
            return sb.ToString();
        }

        private static string BuildParagraphProperties(string alignment, int spacingBeforePt = 0, int spacingAfterPt = 0, int indentLeftPt = 0, int indentRightPt = 0)
        {
            bool hasContent = !string.IsNullOrEmpty(alignment) || spacingBeforePt > 0 || spacingAfterPt > 0 || indentLeftPt > 0 || indentRightPt > 0;
            if (!hasContent)
                return string.Empty;

            var sb = new StringBuilder("<w:pPr>");
            if (!string.IsNullOrEmpty(alignment))
                sb.Append($"<w:jc w:val=\"{alignment}\"/>");
            if (spacingBeforePt > 0 || spacingAfterPt > 0)
            {
                sb.Append("<w:spacing");
                if (spacingBeforePt > 0) sb.Append($" w:before=\"{spacingBeforePt * 20}\"");
                if (spacingAfterPt > 0) sb.Append($" w:after=\"{spacingAfterPt * 20}\"");
                sb.Append("/>");
            }
            if (indentLeftPt > 0 || indentRightPt > 0)
            {
                sb.Append("<w:ind");
                if (indentLeftPt > 0) sb.Append($" w:left=\"{indentLeftPt * 20}\"");
                if (indentRightPt > 0) sb.Append($" w:right=\"{indentRightPt * 20}\"");
                sb.Append("/>");
            }
            sb.Append("</w:pPr>");
            return sb.ToString();
        }

        internal static string EscapeXml(string text)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;
            return text
                .Replace("&", "&amp;")
                .Replace("<", "&lt;")
                .Replace(">", "&gt;")
                .Replace("\"", "&quot;")
                .Replace("'", "&apos;");
        }

        // ── Explicit IDocument implementation ─────────────────────────
        // The concrete methods above return DocxDocument for fluent chaining.
        // These explicit forwarders satisfy the IDocument interface contract.

        IDocument IDocument.SetPageMargins(int top, int bottom, int left, int right, int header, int footer)
            => SetPageMargins(top, bottom, left, right, header, footer);

        IDocument IDocument.SetHeader(string text)
            => SetHeader(new DocxHeaderFooter().SetText(text));

        IDocument IDocument.SetFooter(string text, bool includePageNumbers)
        {
            var f = new DocxHeaderFooter().SetText(text ?? "");
            if (includePageNumbers) f.AddPageNumber();
            return SetFooter(f);
        }

        IDocument IDocument.AddPageNumbers() => AddPageNumbers();

        IDocument IDocument.AddParagraph(string text, bool bold, bool italic, int fontSize,
            string fontName, string alignment, int spacingBeforePt, int spacingAfterPt,
            int indentLeftPt, int indentRightPt)
            => AddParagraph(text, bold, italic, fontSize, fontName, alignment,
                spacingBeforePt, spacingAfterPt, indentLeftPt, indentRightPt);

        IDocument IDocument.AddHeading(string text, int level, string bookmarkName)
            => AddHeading(text, level, bookmarkName);

        IDocument IDocument.AddHorizontalRule() => AddHorizontalRule();
        IDocument IDocument.AddPageBreak() => AddPageBreak();
        IDocument IDocument.AddLineBreak() => AddLineBreak();
        IDocument IDocument.AddSectionSeparator(int spacingPt) => AddSectionSeparator(spacingPt);
        IDocument IDocument.AutoSpacing(int after, int before, double multiple) => AutoSpacing(after, before, multiple);
        IDocument IDocument.SetHeadingsStartOnNewPage(int maxLevel) => SetHeadingsStartOnNewPage(maxLevel);
    }
}
