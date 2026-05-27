using System.Text;

namespace DocumentGenerator
{
    /// <summary>
    /// Represents a header or footer section for a DocxDocument.
    /// </summary>
    public class DocxHeaderFooter
    {
        private string _text;
        private string _alignment = "center";
        private bool _bold;
        private int _fontSize;
        private int _spacingBeforePt;
        private int _spacingAfterPt;
        private int _indentLeftPt;
        private int _indentRightPt;
        private string _pageNumberAlignment = "center";

        /// <summary>
        /// The relationship ID assigned when added to the document.
        /// </summary>
        internal string RelationshipId { get; set; }

        /// <summary>
        /// Whether to include automatic page numbering.
        /// </summary>
        internal bool IncludePageNumber { get; set; }

        /// <summary>
        /// Sets the text content.
        /// </summary>
        public DocxHeaderFooter SetText(string text)
        {
            _text = text;
            return this;
        }

        /// <summary>
        /// Sets text alignment (left, center, right).
        /// </summary>
        public DocxHeaderFooter SetAlignment(string alignment)
        {
            _alignment = alignment;
            return this;
        }

        /// <summary>
        /// Sets bold formatting.
        /// </summary>
        public DocxHeaderFooter SetBold(bool bold = true)
        {
            _bold = bold;
            return this;
        }

        /// <summary>
        /// Sets font size in points.
        /// </summary>
        public DocxHeaderFooter SetFontSize(int points)
        {
            _fontSize = points;
            return this;
        }

        /// <summary>
        /// Enables page number display.
        /// </summary>
        public DocxHeaderFooter AddPageNumber()
        {
            IncludePageNumber = true;
            return this;
        }

        /// <summary>
        /// Sets page number alignment (left, center, right). Default is center.
        /// </summary>
        public DocxHeaderFooter SetPageNumberAlignment(string alignment)
        {
            _pageNumberAlignment = alignment;
            return this;
        }

        /// <summary>
        /// Sets spacing before and after the header/footer content in points.
        /// </summary>
        public DocxHeaderFooter SetSpacing(int beforePt = 0, int afterPt = 0)
        {
            _spacingBeforePt = beforePt;
            _spacingAfterPt = afterPt;
            return this;
        }

        /// <summary>
        /// Sets left and right indent (margin) for the header/footer content in points.
        /// </summary>
        public DocxHeaderFooter SetIndent(int leftPt = 0, int rightPt = 0)
        {
            _indentLeftPt = leftPt;
            _indentRightPt = rightPt;
            return this;
        }

        internal string ToHeaderXml()
        {
            return BuildXml("hdr");
        }

        internal string ToFooterXml()
        {
            return BuildXml("ftr");
        }

        private string BuildXml(string elementName)
        {
            var sb = new StringBuilder();
            sb.AppendLine("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
            sb.Append($"<w:{elementName} xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">");

            if (!string.IsNullOrEmpty(_text))
            {
                sb.Append("<w:p>");
                sb.Append("<w:pPr>");
                sb.Append($"<w:jc w:val=\"{_alignment}\"/>");
                if (_spacingBeforePt > 0 || _spacingAfterPt > 0)
                {
                    sb.Append("<w:spacing");
                    if (_spacingBeforePt > 0) sb.Append($" w:before=\"{_spacingBeforePt * 20}\"");
                    if (_spacingAfterPt > 0) sb.Append($" w:after=\"{_spacingAfterPt * 20}\"");
                    sb.Append("/>");
                }
                if (_indentLeftPt > 0 || _indentRightPt > 0)
                {
                    sb.Append("<w:ind");
                    if (_indentLeftPt > 0) sb.Append($" w:left=\"{_indentLeftPt * 20}\"");
                    if (_indentRightPt > 0) sb.Append($" w:right=\"{_indentRightPt * 20}\"");
                    sb.Append("/>");
                }
                sb.Append("</w:pPr>");
                sb.Append("<w:r>");
                if (_bold || _fontSize > 0)
                {
                    sb.Append("<w:rPr>");
                    if (_bold) sb.Append("<w:b/>");
                    if (_fontSize > 0) sb.Append($"<w:sz w:val=\"{_fontSize * 2}\"/>");
                    sb.Append("</w:rPr>");
                }
                sb.Append($"<w:t xml:space=\"preserve\">{DocxDocument.EscapeXml(_text)}</w:t>");
                sb.Append("</w:r>");
                sb.Append("</w:p>");
            }

            if (IncludePageNumber)
            {
                sb.Append("<w:p>");
                sb.Append($"<w:pPr><w:jc w:val=\"{_pageNumberAlignment}\"/></w:pPr>");
                sb.Append("<w:r><w:t xml:space=\"preserve\">Page </w:t></w:r>");
                sb.Append("<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>");
                sb.Append("<w:r><w:instrText xml:space=\"preserve\"> PAGE </w:instrText></w:r>");
                sb.Append("<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>");
                sb.Append("<w:r><w:t>1</w:t></w:r>");
                sb.Append("<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>");
                sb.Append("<w:r><w:t xml:space=\"preserve\"> of </w:t></w:r>");
                sb.Append("<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>");
                sb.Append("<w:r><w:instrText xml:space=\"preserve\"> NUMPAGES </w:instrText></w:r>");
                sb.Append("<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>");
                sb.Append("<w:r><w:t>1</w:t></w:r>");
                sb.Append("<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>");
                sb.Append("</w:p>");
            }

            sb.Append($"</w:{elementName}>");
            return sb.ToString();
        }
    }
}
