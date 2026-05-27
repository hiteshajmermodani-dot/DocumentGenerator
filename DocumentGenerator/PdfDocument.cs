using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace DocumentGenerator
{
    /// <summary>
    /// Generates a PDF 1.4 document without any third-party libraries.
    /// Provides a fluent API that mirrors DocxDocument for easy report generation.
    /// </summary>
    public class PdfDocument : IDocument
    {
        // Page size: Letter 8.5x11 inches in points (1 pt = 1/72 inch)
        private const float PageWidth = 612f;
        private const float PageHeight = 792f;

        private float _marginTop = 72f;
        private float _marginBottom = 72f;
        private float _marginLeft = 72f;
        private float _marginRight = 72f;

        private float ContentWidth => PageWidth - _marginLeft - _marginRight;
        private float UsableHeight => PageHeight - _marginTop - _marginBottom;

        // PDF object tracking
        private readonly List<string> _objects = new List<string>();   // raw object strings
        private readonly List<int> _offsets = new List<int>();         // byte offsets for xref

        // Pages and content
        private readonly List<string> _pageContentStreams = new List<string>();
        private readonly List<StringBuilder> _pages = new List<StringBuilder>();
        internal StringBuilder _current;

        // Pending TOC
        private PdfTableOfContents _pendingToc;
        private const string TocMarker = "%%TOC%%";
        private readonly Dictionary<string, int> _bookmarkPages = new Dictionary<string, int>();  // bookmark -> 0-based page index

        // TOC link annotation rects: list of (bookmarkName, x, y_bottom, width, height) per TOC page
        private readonly List<(string bookmark, float x, float yBottom, float width, float height)> _tocLinkRects
            = new List<(string, float, float, float, float)>();
        private int _tocPageIndex = -1;  // 0-based index of the page that holds the TOC

        // Image tracking: each entry records which page (0-based) an image belongs to,
        // the JPEG bytes, display dimensions, and the XObject name used in the stream
        private sealed class PdfImageInfo
        {
            public int PageIndex;       // 0-based page index
            public byte[] JpegBytes;
            public int PixelWidth;
            public int PixelHeight;
            public float DisplayW;     // points
            public float DisplayH;     // points
            public float X;            // lower-left x in PDF coords
            public float Y;            // lower-left y in PDF coords
            public string Name;        // e.g. "Im0"
        }
        private readonly List<PdfImageInfo> _images = new List<PdfImageInfo>();

        // Page tracking
        private int _currentPageNumber = 1;
        private float _usedHeight = 0f;

        // Font sizes
        private const float FontNormal = 11f;
        private const float FontH1 = 18f;
        private const float FontH2 = 15f;
        private const float FontH3 = 13f;
        private const float LineSpacing = 1.3f;
        private const float ParaSpacingAfter = 6f;

        // Header / footer text
        private string _headerText;
        private string _footerText;
        private bool _pageNumbers;

        public PdfDocument()
        {
            StartNewPage();
        }

        // ── Public API ────────────────────────────────────────────────

        public PdfDocument SetPageMargins(int topPt = 72, int bottomPt = 72, int leftPt = 72, int rightPt = 72,
            int headerPt = 36, int footerPt = 36)
        {
            _marginTop = topPt;
            _marginBottom = bottomPt;
            _marginLeft = leftPt;
            _marginRight = rightPt;
            return this;
        }

        public PdfDocument SetHeader(string text)
        {
            _headerText = text;
            return this;
        }

        public PdfDocument SetFooter(string text, bool includePageNumbers = false)
        {
            _footerText = text;
            _pageNumbers = includePageNumbers;
            return this;
        }

        public PdfDocument AddPageNumbers()
        {
            _pageNumbers = true;
            return this;
        }

        public PdfDocument AddParagraph(string text, bool bold = false, bool italic = false,
            int fontSize = 0, string fontName = null, string alignment = null,
            int spacingBeforePt = 0, int spacingAfterPt = 0, int indentLeftPt = 0, int indentRightPt = 0)
        {
            if (string.IsNullOrEmpty(text))
            {
                AdvanceY(FontNormal * LineSpacing);
                return this;
            }

            float size = fontSize > 0 ? fontSize : FontNormal;
            string font = bold ? "Helvetica-Bold" : "Helvetica";

            float lineH = size * LineSpacing;
            float indent = indentLeftPt;
            float availWidth = ContentWidth - indent - indentRightPt;

            var lines = WrapText(text, availWidth, size);

            AdvanceY(spacingBeforePt);
            foreach (var line in lines)
            {
                EnsureSpace(lineH);
                EmitTextLine(line, font, size, _marginLeft + indent, alignment);
                AdvanceY(lineH);
            }
            // Only advance the after-spacing if there's room; otherwise it will be
            // consumed naturally by the next EnsureSpace call on new page
            float afterSpacing = spacingAfterPt + ParaSpacingAfter;
            if (_usedHeight + afterSpacing <= UsableHeight)
                AdvanceY(afterSpacing);
            return this;
        }

        public PdfDocument AddHeading(string text, int level = 1, string bookmarkName = null)
        {
            if (level < 1) level = 1;
            if (level > 9) level = 9;

            float size = level == 1 ? FontH1 : level == 2 ? FontH2 : FontH3;
            float lineH = size * LineSpacing;
            float spaceBefore = level == 1 ? 16f : level == 2 ? 14f : 12f;

            // Level-1 headings start on a new page
            if (level == 1 && _usedHeight > 0)
                AddPageBreak();

            if (!string.IsNullOrEmpty(bookmarkName))
                _bookmarkPages[bookmarkName] = _currentPageNumber - 1;

            // For sub-headings ensure the space-before + at least one line fits
            if (level > 1)
                EnsureSpace(spaceBefore + lineH * 2f);

            AdvanceY(spaceBefore);
            var lines = WrapText(text, ContentWidth, size);
            foreach (var line in lines)
            {
                EnsureSpace(lineH);
                EmitTextLine(line, "Helvetica-Bold", size, _marginLeft, null);
                AdvanceY(lineH);
            }
            AdvanceY(8f);
            return this;
        }

        public PdfDocument AddTable(PdfTable table)
        {
            table.Render(this);
            // Ensure a visible gap after every table before next heading/paragraph
            EnsureSpace(14f);
            AdvanceY(14f);
            return this;
        }

        public PdfDocument AddPageBreak()
        {
            FlushPage();
            StartNewPage();
            return this;
        }

        public PdfDocument AddHorizontalRule()
        {
            EnsureSpace(20f);
            float x1 = _marginLeft;
            float x2 = PageWidth - _marginRight;
            float y = PdfY(_usedHeight + 6f);
            _current.AppendLine($"0.5 w");
            _current.AppendLine($"{x1:F1} {y:F1} m {x2:F1} {y:F1} l S");
            AdvanceY(20f);
            return this;
        }

        public PdfDocument AddSectionSeparator(int spacingPt = 24)
        {
            AdvanceY(spacingPt);
            return this;
        }

        public PdfDocument AddLineBreak()
        {
            AdvanceY(FontNormal * LineSpacing);
            return this;
        }

        public PdfDocument AutoSpacing(int spacingAfterPt = 8, int spacingBeforePt = 0, double lineSpacingMultiple = 1.15)
        {
            // Absorbed — line spacing is built in
            return this;
        }

        public PdfDocument SetHeadingsStartOnNewPage(int maxLevel = 1)
        {
            // Already implemented: level-1 headings always start on new page
            return this;
        }

        /// <summary>
        /// Embeds a JPEG image into the PDF at the current position.
        /// widthPt / heightPt are the display dimensions in points (72 pts = 1 inch).
        /// Pass 0 for widthPt to use the full content width, maintaining aspect ratio.
        /// </summary>
        public PdfDocument AddImage(byte[] jpegBytes, float widthPt = 0f, float heightPt = 0f,
            string alignment = "left")
        {
            if (jpegBytes == null || jpegBytes.Length == 0) return this;

            // Parse pixel dims from JPEG SOF marker so we can compute aspect ratio
            (int px, int py) = ReadJpegDimensions(jpegBytes);
            if (px <= 0) px = 1;
            if (py <= 0) py = 1;

            if (widthPt <= 0f && heightPt <= 0f)
            {
                widthPt = ContentWidth;
                heightPt = ContentWidth * py / px;
            }
            else if (widthPt <= 0f)
            {
                widthPt = heightPt * px / py;
            }
            else if (heightPt <= 0f)
            {
                heightPt = widthPt * py / px;
            }

            // Clamp to page width
            if (widthPt > ContentWidth)
            {
                heightPt = heightPt * ContentWidth / widthPt;
                widthPt = ContentWidth;
            }

            EnsureSpace(heightPt + 6f);

            float imgX = _marginLeft;
            if (alignment == "center")
                imgX = _marginLeft + (ContentWidth - widthPt) / 2f;
            else if (alignment == "right")
                imgX = PageWidth - _marginRight - widthPt;

            // PDF Y is bottom-up; image bottom-left is below the current top position
            float imgY = PdfY(_usedHeight + heightPt);

            string imgName = $"Im{_images.Count}";
            int pageIdx = _currentPageNumber - 1;

            _images.Add(new PdfImageInfo
            {
                PageIndex = pageIdx,
                JpegBytes = jpegBytes,
                PixelWidth = px,
                PixelHeight = py,
                DisplayW = widthPt,
                DisplayH = heightPt,
                X = imgX,
                Y = imgY,
                Name = imgName
            });

            // Emit the Do command into the current page stream
            _current.AppendLine($"q {widthPt:F2} 0 0 {heightPt:F2} {imgX:F2} {imgY:F2} cm /{imgName} Do Q");

            AdvanceY(heightPt + 6f);
            return this;
        }

        /// <summary>Reads image width/height from a JPEG SOF0/SOF2 marker.</summary>
        private static (int width, int height) ReadJpegDimensions(byte[] data)
        {
            int i = 2;  // skip SOI FF D8
            while (i < data.Length - 8)
            {
                if (data[i] != 0xFF) break;
                byte marker = data[i + 1];
                int segLen = (data[i + 2] << 8) | data[i + 3];

                // SOF markers: C0..C3, C5..C7, C9..CB, CD..CF
                if (marker >= 0xC0 && marker <= 0xCF && marker != 0xC4 && marker != 0xC8 && marker != 0xCC)
                {
                    int height = (data[i + 5] << 8) | data[i + 6];
                    int width  = (data[i + 7] << 8) | data[i + 8];
                    return (width, height);
                }
                i += 2 + segLen;
            }
            return (0, 0);
        }

        public PdfDocument AddTableOfContents(PdfTableOfContents toc)
        {
            _pendingToc = toc;
            _tocPageIndex = _currentPageNumber - 1;  // 0-based index of current page
            _current.AppendLine($"%{TocMarker}");
            return this;
        }

        // ── Internal rendering helpers ────────────────────────────────

        internal void EnsureSpace(float needed)
        {
            // If the item itself is larger than a full page, just start a fresh page
            // (avoids an infinite flush loop)
            if (needed >= UsableHeight)
            {
                FlushPage();
                StartNewPage();
                return;
            }
            if (_usedHeight + needed > UsableHeight)
            {
                FlushPage();
                StartNewPage();
            }
        }

        internal void EmitTextLine(string text, string font, float size, float x, string alignment)
        {
            if (string.IsNullOrEmpty(text)) return;
            float y = PdfY(_usedHeight);
            string safe = EscapePdf(text);

            // Right-align: measure approximate width
            if (alignment == "right")
            {
                float w = MeasureText(text, size);
                x = PageWidth - _marginRight - w;
            }
            else if (alignment == "center")
            {
                float w = MeasureText(text, size);
                x = _marginLeft + (ContentWidth - w) / 2f;
            }

            _current.AppendLine($"BT /{font} {size:F1} Tf {x:F1} {y:F1} Td ({safe}) Tj ET");
        }

        internal void EmitTextAt(string text, string font, float size, float x, float absY)
        {
            if (string.IsNullOrEmpty(text)) return;
            string safe = EscapePdf(text);
            _current.AppendLine($"BT /{font} {size:F1} Tf {x:F1} {absY:F1} Td ({safe}) Tj ET");
        }

        internal void EmitLine(float x1, float absY1, float x2, float absY2)
        {
            _current.AppendLine($"{x1:F1} {absY1:F1} m {x2:F1} {absY2:F1} l S");
        }

        internal void EmitRect(float x, float absY, float w, float h)
        {
            _current.AppendLine($"{x:F1} {absY:F1} {w:F1} {h:F1} re S");
        }

        internal void AdvanceY(float pts)
        {
            if (pts <= 0f) return;
            _usedHeight += pts;
            // Clamp: if we've drifted past the bottom margin, cap it so PdfY never goes
            // below the margin. The next EnsureSpace call will start a fresh page.
            if (_usedHeight > UsableHeight)
                _usedHeight = UsableHeight;
        }

        internal float CurrentY => _usedHeight;
        internal float MarginLeft => _marginLeft;
        internal float MarginRight => _marginRight;
        internal float PageW => PageWidth;
        internal float ContentW => ContentWidth;
        internal float UsableH => UsableHeight;
        internal int CurrentPageNumber => _currentPageNumber;

        internal float PdfY(float usedFromTop)
        {
            return PageHeight - _marginTop - usedFromTop;
        }

        // ── Page management ───────────────────────────────────────────

        private void StartNewPage()
        {
            _current = new StringBuilder();
            // Set line width for rules/table borders
            _current.AppendLine("0.5 w");
            _usedHeight = 0f;
        }

        private void FlushPage()
        {
            // Bake header/footer into stream before saving
            var stream = new StringBuilder();
            stream.Append(_current);
            AppendHeaderFooter(stream, _currentPageNumber);
            _pageContentStreams.Add(stream.ToString());
            _currentPageNumber++;
        }

        private void AppendHeaderFooter(StringBuilder stream, int pageNum)
        {
            if (!string.IsNullOrEmpty(_headerText))
            {
                float y = PageHeight - _marginTop / 2f;
                float w = MeasureText(_headerText, 9f);
                float x = _marginLeft + (ContentWidth - w) / 2f;
                stream.AppendLine($"BT /Helvetica 9 Tf {x:F1} {y:F1} Td ({EscapePdf(_headerText)}) Tj ET");
            }

            string footerLine = _footerText ?? "";
            if (_pageNumbers)
            {
                string pageStr = _footerText != null ? $"{_footerText}    Page {pageNum}" : $"Page {pageNum}";
                footerLine = pageStr;
            }

            if (!string.IsNullOrEmpty(footerLine))
            {
                float y = _marginBottom / 2f;
                float w = MeasureText(footerLine, 9f);
                float x = _marginLeft + (ContentWidth - w) / 2f;
                stream.AppendLine($"BT /Helvetica 9 Tf {x:F1} {y:F1} Td ({EscapePdf(footerLine)}) Tj ET");
            }
        }

        // ── Save ──────────────────────────────────────────────────────

        public void Save(string filePath)
        {
            // Flush last page
            FlushPage();

            // Render TOC content into the placeholder page
            if (_pendingToc != null && _tocPageIndex >= 0 && _tocPageIndex < _pageContentStreams.Count)
            {
                _pendingToc.ComputeSerialNumbers();
                var tocResult = BuildTocPageContent(_pendingToc, _tocPageIndex + 1);
                _pageContentStreams[_tocPageIndex] = tocResult.content;
                _tocLinkRects.Clear();
                _tocLinkRects.AddRange(tocResult.linkRects);
            }

            using var ms = new MemoryStream();
            using var w = new StreamWriter(ms, new UTF8Encoding(false));

            w.WriteLine("%PDF-1.4");
            w.WriteLine("%\xE2\xE3\xCF\xD3");
            w.Flush();

            var offsets = new List<long>();

            int nPages = _pageContentStreams.Count;

            // ── Object layout ─────────────────────────────────────────────────────────
            // 1                           catalog
            // 2                           page tree
            // 3 .. 2+nPages               content streams (one per page)
            // 3+nPages .. 2+2*nPages      page objects
            // 3+2*nPages                  font Helvetica
            // 4+2*nPages                  font Helvetica-Bold
            // 5+2*nPages .. +nDests-1     named destination objects
            // +nDests .. +nDests+nAnnots-1 link annotation objects
            // +nDests+nAnnots ..          image XObject streams (one per image)

            int contentBase = 3;
            int pageBase    = contentBase + nPages;
            int fontHelv    = pageBase + nPages;
            int fontBold    = fontHelv + 1;

            var bookmarkList = new List<(string name, int pageIndex)>();
            foreach (var kv in _bookmarkPages)
                bookmarkList.Add((kv.Key, kv.Value));

            int nDests    = bookmarkList.Count;
            int destBase  = fontBold + 1;
            int nAnnots   = _tocLinkRects.Count;
            int annotBase = destBase + nDests;
            int imgBase   = annotBase + nAnnots;          // image XObjects start here
            int nImgs     = _images.Count;
            int lastObj   = imgBase + nImgs - 1;
            int totalObjects = (nImgs > 0) ? lastObj
                             : (nAnnots > 0) ? annotBase + nAnnots - 1
                             : (nDests > 0)  ? destBase + nDests - 1
                             : fontBold;

            // Pre-build per-page XObject name→object-number map
            // e.g. page 3 has Im0 at imgBase+0  and Im2 at imgBase+2
            var pageXObjects = new Dictionary<int, List<(string name, int obj)>>();
            for (int ii = 0; ii < _images.Count; ii++)
            {
                int pg = _images[ii].PageIndex;
                if (!pageXObjects.ContainsKey(pg))
                    pageXObjects[pg] = new List<(string, int)>();
                pageXObjects[pg].Add((_images[ii].Name, imgBase + ii));
            }

            // Bookmark → destination obj
            var bookmarkToDestObj = new Dictionary<string, int>();
            for (int i = 0; i < bookmarkList.Count; i++)
                bookmarkToDestObj[bookmarkList[i].name] = destBase + i;

            // ── Obj 1: Catalog ────────────────────────────────────────────────────────
            w.Flush();
            offsets.Add(ms.Position);
            string catalogExtra = "";
            if (nDests > 0)
            {
                var namePairs = new StringBuilder();
                for (int i = 0; i < bookmarkList.Count; i++)
                {
                    if (i > 0) namePairs.Append(" ");
                    namePairs.Append($"({EscapePdf(bookmarkList[i].name)}) {destBase + i} 0 R");
                }
                catalogExtra = $" /Names << /Dests << /Names [{namePairs}] >> >>";
            }
            w.WriteLine("1 0 obj");
            w.WriteLine($"<< /Type /Catalog /Pages 2 0 R{catalogExtra} >>");
            w.WriteLine("endobj");
            w.Flush();

            // ── Obj 2: Page tree ──────────────────────────────────────────────────────
            var pageRefs = new StringBuilder();
            for (int i = 0; i < nPages; i++)
                pageRefs.Append($"{pageBase + i} 0 R ");
            w.Flush();
            offsets.Add(ms.Position);
            w.WriteLine("2 0 obj");
            w.WriteLine($"<< /Type /Pages /Kids [{pageRefs}] /Count {nPages} >>");
            w.WriteLine("endobj");
            w.Flush();

            // ── Content streams ───────────────────────────────────────────────────────
            for (int i = 0; i < nPages; i++)
            {
                byte[] streamBytes = Encoding.GetEncoding("iso-8859-1").GetBytes(_pageContentStreams[i]);
                w.Flush();
                offsets.Add(ms.Position);
                w.WriteLine($"{contentBase + i} 0 obj");
                w.WriteLine($"<< /Length {streamBytes.Length} >>");
                w.WriteLine("stream");
                w.Flush();
                ms.Write(streamBytes, 0, streamBytes.Length);
                w.WriteLine();
                w.WriteLine("endstream");
                w.WriteLine("endobj");
                w.Flush();
            }

            // ── Page objects ──────────────────────────────────────────────────────────
            var tocAnnotRefs = new StringBuilder();
            if (nAnnots > 0)
                for (int a = 0; a < nAnnots; a++)
                    tocAnnotRefs.Append($"{annotBase + a} 0 R ");

            for (int i = 0; i < nPages; i++)
            {
                string annotEntry = (i == _tocPageIndex && tocAnnotRefs.Length > 0)
                    ? $" /Annots [{tocAnnotRefs}]" : "";

                // Build XObject dict for this page
                string xobjEntry = "";
                if (pageXObjects.TryGetValue(i, out var xobjs) && xobjs.Count > 0)
                {
                    var xsb = new StringBuilder();
                    foreach (var (xname, xobj) in xobjs)
                        xsb.Append($" /{xname} {xobj} 0 R");
                    xobjEntry = $" /XObject <<{xsb} >>";
                }

                w.Flush();
                offsets.Add(ms.Position);
                w.WriteLine($"{pageBase + i} 0 obj");
                w.WriteLine($"<< /Type /Page /Parent 2 0 R " +
                    $"/MediaBox [0 0 {PageWidth:F0} {PageHeight:F0}] " +
                    $"/Contents {contentBase + i} 0 R " +
                    $"/Resources << /Font << /Helvetica {fontHelv} 0 R /Helvetica-Bold {fontBold} 0 R >>{xobjEntry} >>" +
                    $"{annotEntry} >>");
                w.WriteLine("endobj");
                w.Flush();
            }

            // ── Font objects ──────────────────────────────────────────────────────────
            w.Flush();
            offsets.Add(ms.Position);
            w.WriteLine($"{fontHelv} 0 obj");
            w.WriteLine("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>");
            w.WriteLine("endobj");
            w.Flush();

            w.Flush();
            offsets.Add(ms.Position);
            w.WriteLine($"{fontBold} 0 obj");
            w.WriteLine("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >>");
            w.WriteLine("endobj");
            w.Flush();

            // ── Named destination objects ─────────────────────────────────────────────
            for (int i = 0; i < bookmarkList.Count; i++)
            {
                var (name, pageIndex) = bookmarkList[i];
                int targetPageObj = (pageIndex >= 0 && pageIndex < nPages) ? pageBase + pageIndex : pageBase;
                w.Flush();
                offsets.Add(ms.Position);
                w.WriteLine($"{destBase + i} 0 obj");
                w.WriteLine($"[{targetPageObj} 0 R /XYZ null null null]");
                w.WriteLine("endobj");
                w.Flush();
            }

            // ── Link annotation objects ───────────────────────────────────────────────
            for (int a = 0; a < nAnnots; a++)
            {
                var (bmName, lx, ly, lw, lh) = _tocLinkRects[a];
                string destRef = bookmarkToDestObj.TryGetValue(bmName, out int dObj)
                    ? $"{dObj} 0 R" : "null";
                w.Flush();
                offsets.Add(ms.Position);
                w.WriteLine($"{annotBase + a} 0 obj");
                w.WriteLine($"<< /Type /Annot /Subtype /Link " +
                    $"/Rect [{lx:F1} {ly:F1} {(lx + lw):F1} {(ly + lh):F1}] " +
                    $"/Border [0 0 0] /Dest {destRef} >>");
                w.WriteLine("endobj");
                w.Flush();
            }

            // ── Image XObject streams ─────────────────────────────────────────────────
            for (int ii = 0; ii < _images.Count; ii++)
            {
                var img = _images[ii];
                w.Flush();
                offsets.Add(ms.Position);
                w.WriteLine($"{imgBase + ii} 0 obj");
                w.WriteLine($"<< /Type /XObject /Subtype /Image " +
                    $"/Width {img.PixelWidth} /Height {img.PixelHeight} " +
                    $"/ColorSpace /DeviceRGB /BitsPerComponent 8 " +
                    $"/Filter /DCTDecode /Length {img.JpegBytes.Length} >>");
                w.WriteLine("stream");
                w.Flush();
                ms.Write(img.JpegBytes, 0, img.JpegBytes.Length);
                w.WriteLine();
                w.WriteLine("endstream");
                w.WriteLine("endobj");
                w.Flush();
            }

            // ── Cross-reference table ─────────────────────────────────────────────────
            w.Flush();
            long xrefOffset = ms.Position;
            w.WriteLine("xref");
            w.WriteLine($"0 {totalObjects + 1}");
            w.WriteLine("0000000000 65535 f ");
            foreach (long off in offsets)
                w.WriteLine($"{off:D10} 00000 n ");
            w.Flush();

            w.WriteLine("trailer");
            w.WriteLine($"<< /Size {totalObjects + 1} /Root 1 0 R >>");
            w.WriteLine("startxref");
            w.WriteLine(xrefOffset.ToString());
            w.WriteLine("%%EOF");
            w.Flush();

            File.WriteAllBytes(filePath, ms.ToArray());
        }

        private void WriteObj(StreamWriter w, MemoryStream ms, List<long> offsets, int num, string dict)
        {
            w.Flush();
            offsets.Add(ms.Position);
            w.WriteLine($"{num} 0 obj");
            w.WriteLine(dict);
            w.WriteLine("endobj");
            w.Flush();
        }

        private void WriteStreamObj(StreamWriter w, MemoryStream ms, List<long> offsets, int num, string dict, byte[] data)
        {
            w.Flush();
            offsets.Add(ms.Position);
            w.WriteLine($"{num} 0 obj");
            w.WriteLine(dict);
            w.WriteLine("stream");
            w.Flush();
            ms.Write(data, 0, data.Length);
            w.WriteLine();
            w.WriteLine("endstream");
            w.WriteLine("endobj");
            w.Flush();
        }

        // ── TOC rendering ─────────────────────────────────────────────

        private (string content, List<(string bookmark, float x, float yBottom, float width, float height)> linkRects)
            BuildTocPageContent(PdfTableOfContents toc, int pageNumber)
        {
            var sb = new StringBuilder();
            sb.AppendLine("0.5 w");
            var links = new List<(string, float, float, float, float)>();

            float y = 0f;

            // Title
            float titleSize = FontH1;
            float titleY = PdfYStatic(y);
            string titleText = toc.Title ?? "Table of Contents";
            float titleW = MeasureText(titleText, titleSize);
            float titleX = _marginLeft + (ContentWidth - titleW) / 2f;
            sb.AppendLine($"BT /Helvetica-Bold {titleSize:F1} Tf {titleX:F1} {titleY:F1} Td ({EscapePdf(titleText)}) Tj ET");
            y += titleSize * LineSpacing + 10f;

            foreach (var entry in toc.Entries)
            {
                int level = Math.Max(1, Math.Min(9, entry.Level));
                float size = level == 1 ? 12f : 11f;
                string font = level == 1 ? "Helvetica-Bold" : "Helvetica";
                float indent = (level - 1) * 18f;
                float lineH = size * LineSpacing + 2f;

                string displayText = toc.ShowNumbers && !string.IsNullOrEmpty(entry.SerialNumber)
                    ? entry.SerialNumber + "  " + entry.Title
                    : entry.Title;

                int cachedPage = 1;
                if (_bookmarkPages.TryGetValue(entry.BookmarkName ?? "", out int bp))
                    cachedPage = bp + 1;  // convert 0-based to 1-based for display

                string pageStr = cachedPage.ToString();
                float entryY = PdfYStatic(y);
                float textX = _marginLeft + indent;
                float pageNumX = PageWidth - _marginRight - MeasureText(pageStr, size);

                // Dots
                float dotsStart = textX + MeasureText(displayText, size) + 4f;
                float dotsEnd = pageNumX - 4f;
                float dotStep = MeasureText(".", size) + 1f;
                var dots = new StringBuilder();
                for (float dx = dotsStart; dx < dotsEnd; dx += dotStep) dots.Append(".");

                // Draw entry text in blue for link style
                string r = "0", g = "0", b = "0";
                if (toc.LinkStyle) { r = "0.07"; g = "0.36"; b = "0.80"; }

                if (toc.LinkStyle)
                    sb.AppendLine($"{r} {g} {b} rg");

                sb.AppendLine($"BT /{font} {size:F1} Tf {textX:F1} {entryY:F1} Td ({EscapePdf(displayText)}) Tj ET");
                if (dots.Length > 0)
                    sb.AppendLine($"BT /Helvetica {size:F1} Tf {dotsStart:F1} {entryY:F1} Td ({EscapePdf(dots.ToString())}) Tj ET");
                sb.AppendLine($"BT /Helvetica {size:F1} Tf {pageNumX:F1} {entryY:F1} Td ({pageStr}) Tj ET");

                if (toc.LinkStyle)
                    sb.AppendLine("0 0 0 rg");

                // Record link annotation rect covering the full line
                if (!string.IsNullOrEmpty(entry.BookmarkName))
                {
                    float rectX = textX;
                    float rectW = (PageWidth - _marginRight) - textX;
                    float rectH = size + 4f;
                    float rectBottom = entryY - 2f;
                    links.Add((entry.BookmarkName, rectX, rectBottom, rectW, rectH));
                }

                y += lineH;
            }

            AppendHeaderFooter(sb, pageNumber);
            return (sb.ToString(), links);
        }

        private float PdfYStatic(float usedFromTop) => PageHeight - _marginTop - usedFromTop;

        // ── Text utilities ────────────────────────────────────────────

        internal List<string> WrapText(string text, float maxWidth, float fontSize)
        {
            var result = new List<string>();
            if (string.IsNullOrEmpty(text)) { result.Add(""); return result; }

            // Average char width for Helvetica ≈ 0.55 * fontSize
            float charW = fontSize * 0.55f;
            int charsPerLine = Math.Max(1, (int)(maxWidth / charW));

            string[] words = text.Split(' ');
            var line = new StringBuilder();
            foreach (var word in words)
            {
                if (line.Length == 0)
                {
                    line.Append(word);
                }
                else if (line.Length + 1 + word.Length <= charsPerLine)
                {
                    line.Append(' ');
                    line.Append(word);
                }
                else
                {
                    result.Add(line.ToString());
                    line.Clear();
                    line.Append(word);
                }
            }
            if (line.Length > 0) result.Add(line.ToString());
            return result;
        }

        internal float MeasureText(string text, float fontSize)
        {
            if (string.IsNullOrEmpty(text)) return 0f;
            return text.Length * fontSize * 0.55f;
        }

        internal static string EscapePdf(string text)
        {
            if (string.IsNullOrEmpty(text)) return "";
            return text
                .Replace("\\", "\\\\")
                .Replace("(", "\\(")
                .Replace(")", "\\)")
                .Replace("\r", "")
                .Replace("\n", " ");
        }

        // ── Explicit IDocument implementation ─────────────────────────
        // The concrete methods above return PdfDocument for fluent chaining.
        // These explicit forwarders satisfy the IDocument interface contract.

        IDocument IDocument.SetPageMargins(int top, int bottom, int left, int right, int header, int footer)
            => SetPageMargins(top, bottom, left, right, header, footer);

        IDocument IDocument.SetHeader(string text) => SetHeader(text);

        IDocument IDocument.SetFooter(string text, bool includePageNumbers)
            => SetFooter(text, includePageNumbers);

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
