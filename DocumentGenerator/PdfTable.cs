using System;
using System.Collections.Generic;

namespace DocumentGenerator
{
    /// <summary>
    /// Builds a table for inclusion in a PdfDocument.
    /// Mirrors the DocxTable API for use in shared report-building code.
    /// </summary>
    public class PdfTable
    {
        private readonly List<string[]> _rows = new List<string[]>();
        private readonly List<bool> _isHeader = new List<bool>();
        private string _borderColor = "4472C4";
        private string _headerBgColor = "4472C4";
        private bool _drawBorders = true;

        public int RowCount => _rows.Count;

        /// <summary>Sets border visibility and color (hex RGB, e.g. "4472C4").</summary>
        public PdfTable SetBorders(string style = "single", int size = 4, string color = "4472C4")
        {
            _borderColor = color;
            // Use the same color as header background, but only if it's not pure black
            // (pure black headers make text invisible)
            _headerBgColor = (color == "000000") ? "4472C4" : color;
            _drawBorders = true;
            return this;
        }

        public PdfTable SetAlignment(string alignment) => this;
        public PdfTable SetIndent(int indentPt) => this;
        public PdfTable SetCellPadding(int topPt = 4, int bottomPt = 4, int leftPt = 4, int rightPt = 4) => this;

        /// <summary>Adds a bold header row.</summary>
        public PdfTable AddHeaderRow(params string[] cells)
        {
            _rows.Add(cells);
            _isHeader.Add(true);
            return this;
        }

        /// <summary>Adds a data row.</summary>
        public PdfTable AddRow(params string[] cells)
        {
            _rows.Add(cells);
            _isHeader.Add(false);
            return this;
        }

        /// <summary>Renders the table into the PdfDocument at the current position.</summary>
        internal void Render(PdfDocument doc)
        {
            if (_rows.Count == 0) return;

            int cols = _rows[0].Length;
            float colWidth = doc.ContentW / cols;
            float cellPad = 3f;
            float rowHeight = 16f;

            for (int r = 0; r < _rows.Count; r++)
            {
                // EnsureSpace FIRST — this may flush to a new page
                doc.EnsureSpace(rowHeight);

                // Recalculate rowY AFTER EnsureSpace so coordinates are correct on current page
                float rowTopPdf = doc.PdfY(doc.CurrentY);
                float rowBottomPdf = rowTopPdf - rowHeight;
                bool header = _isHeader[r];

                if (header)
                {
                    float[] rgb = HexToRgb(_headerBgColor);
                    doc._current.AppendLine($"{rgb[0]:F2} {rgb[1]:F2} {rgb[2]:F2} rg");
                    doc._current.AppendLine($"{doc.MarginLeft:F1} {rowBottomPdf:F1} {doc.ContentW:F1} {rowHeight:F1} re f");
                    doc._current.AppendLine("0 0 0 rg");
                }

                for (int c = 0; c < cols && c < _rows[r].Length; c++)
                {
                    float cellX = doc.MarginLeft + c * colWidth;
                    float textY = rowBottomPdf + cellPad + 2f;
                    string font = header ? "Helvetica-Bold" : "Helvetica";
                    float fontSize = 9f;

                    // Use white text on header rows for readability
                    if (header)
                        doc._current.AppendLine("1 1 1 rg");

                    string cellText = TruncateText(_rows[r][c] ?? "", colWidth - cellPad * 2, fontSize);
                    doc.EmitTextAt(cellText, font, fontSize, cellX + cellPad, textY);

                    if (header)
                        doc._current.AppendLine("0 0 0 rg");  // reset to black for borders/data rows

                    if (_drawBorders)
                        doc.EmitRect(cellX, rowBottomPdf, colWidth, rowHeight);
                }

                // Advance AFTER drawing so next row/element starts below this row
                doc.AdvanceY(rowHeight);
            }
            // Spacing after table — only if room remains; otherwise next EnsureSpace handles it
            if (doc.CurrentY + 8f <= doc.UsableH)
                doc.AdvanceY(8f);
        }

        private static float[] HexToRgb(string hex)
        {
            if (hex.Length < 6) return new float[] { 0f, 0f, 0f };
            try
            {
                int r = Convert.ToInt32(hex.Substring(0, 2), 16);
                int g = Convert.ToInt32(hex.Substring(2, 2), 16);
                int b = Convert.ToInt32(hex.Substring(4, 2), 16);
                return new float[] { r / 255f, g / 255f, b / 255f };
            }
            catch
            {
                return new float[] { 0.8f, 0.8f, 0.8f };
            }
        }

        private string TruncateText(string text, float maxWidth, float fontSize)
        {
            float charW = fontSize * 0.55f;
            int maxChars = Math.Max(1, (int)(maxWidth / charW));
            if (text.Length <= maxChars) return text;
            return text.Substring(0, Math.Max(0, maxChars - 2)) + "..";
        }
    }
}
