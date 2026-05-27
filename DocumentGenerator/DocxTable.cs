using System.Collections.Generic;
using System.Text;

namespace DocumentGenerator
{
    /// <summary>
    /// Builds a table for inclusion in a DocxDocument.
    /// </summary>
    public class DocxTable
    {
        private readonly List<DocxTableRow> _rows = new List<DocxTableRow>();
        private string _borderStyle = "single";
        private int _borderSize = 4;
        private string _borderColor = "000000";
        private string _alignment;
        private int _indentTwips;
        private int _cellPaddingTopPt;
        private int _cellPaddingBottomPt;
        private int _cellPaddingLeftPt;
        private int _cellPaddingRightPt;

        /// <summary>
        /// Sets the table border style.
        /// </summary>
        public DocxTable SetBorders(string style = "single", int size = 4, string color = "000000")
        {
            _borderStyle = style;
            _borderSize = size;
            _borderColor = color;
            return this;
        }

        /// <summary>
        /// Sets the table alignment (left, center, right).
        /// </summary>
        public DocxTable SetAlignment(string alignment)
        {
            _alignment = alignment;
            return this;
        }

        /// <summary>
        /// Sets the table left indent in points.
        /// </summary>
        public DocxTable SetIndent(int indentPt)
        {
            _indentTwips = indentPt * 20;
            return this;
        }

        /// <summary>
        /// Sets cell padding (margin) for all cells in points.
        /// </summary>
        public DocxTable SetCellPadding(int topPt = 0, int bottomPt = 0, int leftPt = 0, int rightPt = 0)
        {
            _cellPaddingTopPt = topPt;
            _cellPaddingBottomPt = bottomPt;
            _cellPaddingLeftPt = leftPt;
            _cellPaddingRightPt = rightPt;
            return this;
        }

        /// <summary>
        /// Returns the total number of rows (including header row) for height estimation.
        /// </summary>
        public int RowCount => _rows.Count;

        /// <summary>
        /// Adds a header row with bold text.
        /// </summary>
        public DocxTable AddHeaderRow(params string[] cells)
        {
            _rows.Add(new DocxTableRow(cells, true));
            return this;
        }

        /// <summary>
        /// Adds a data row.
        /// </summary>
        public DocxTable AddRow(params string[] cells)
        {
            _rows.Add(new DocxTableRow(cells, false));
            return this;
        }

        internal string ToXml()
        {
            var sb = new StringBuilder();
            sb.Append("<w:tbl>");
            sb.Append("<w:tblPr>");

            // Borders
            sb.Append("<w:tblBorders>");
            string border = $"w:val=\"{_borderStyle}\" w:sz=\"{_borderSize}\" w:space=\"0\" w:color=\"{_borderColor}\"";
            sb.Append($"<w:top {border}/>");
            sb.Append($"<w:left {border}/>");
            sb.Append($"<w:bottom {border}/>");
            sb.Append($"<w:right {border}/>");
            sb.Append($"<w:insideH {border}/>");
            sb.Append($"<w:insideV {border}/>");
            sb.Append("</w:tblBorders>");

            sb.Append("<w:tblW w:w=\"5000\" w:type=\"pct\"/>");

            // Alignment
            if (!string.IsNullOrEmpty(_alignment))
                sb.Append($"<w:jc w:val=\"{_alignment}\"/>");

            // Indent
            if (_indentTwips > 0)
                sb.Append($"<w:tblInd w:w=\"{_indentTwips}\" w:type=\"dxa\"/>");

            // Cell margins (padding)
            if (_cellPaddingTopPt > 0 || _cellPaddingBottomPt > 0 || _cellPaddingLeftPt > 0 || _cellPaddingRightPt > 0)
            {
                sb.Append("<w:tblCellMar>");
                if (_cellPaddingTopPt > 0) sb.Append($"<w:top w:w=\"{_cellPaddingTopPt * 20}\" w:type=\"dxa\"/>");
                if (_cellPaddingBottomPt > 0) sb.Append($"<w:bottom w:w=\"{_cellPaddingBottomPt * 20}\" w:type=\"dxa\"/>");
                if (_cellPaddingLeftPt > 0) sb.Append($"<w:left w:w=\"{_cellPaddingLeftPt * 20}\" w:type=\"dxa\"/>");
                if (_cellPaddingRightPt > 0) sb.Append($"<w:right w:w=\"{_cellPaddingRightPt * 20}\" w:type=\"dxa\"/>");
                sb.Append("</w:tblCellMar>");
            }

            sb.Append("</w:tblPr>");

            foreach (var row in _rows)
                sb.Append(row.ToXml());

            sb.Append("</w:tbl>");
            return sb.ToString();
        }

        private class DocxTableRow
        {
            private readonly string[] _cells;
            private readonly bool _isHeader;

            public DocxTableRow(string[] cells, bool isHeader)
            {
                _cells = cells;
                _isHeader = isHeader;
            }

            public string ToXml()
            {
                var sb = new StringBuilder();
                sb.Append("<w:tr>");

                if (_isHeader)
                    sb.Append("<w:trPr><w:tblHeader/></w:trPr>");

                foreach (var cell in _cells)
                {
                    sb.Append("<w:tc>");
                    sb.Append("<w:p><w:r>");
                    if (_isHeader)
                        sb.Append("<w:rPr><w:b/></w:rPr>");
                    sb.Append($"<w:t xml:space=\"preserve\">{DocxDocument.EscapeXml(cell)}</w:t>");
                    sb.Append("</w:r></w:p>");
                    sb.Append("</w:tc>");
                }

                sb.Append("</w:tr>");
                return sb.ToString();
            }
        }
    }
}
