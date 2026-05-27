using System.Collections.Generic;
using System.Text;

namespace DocumentGenerator
{
    /// <summary>
    /// Represents a single entry in the Table of Contents.
    /// </summary>
    public class TocEntry
    {
        public string Title { get; set; }
        public int Level { get; set; }
        public string BookmarkName { get; set; }
        public string SerialNumber { get; internal set; }

        public TocEntry(string title, int level, string bookmarkName)
        {
            Title = title;
            Level = level;
            BookmarkName = bookmarkName;
        }
    }

    /// <summary>
    /// Builds a structured Table of Contents with tree-style numbering and indentation.
    /// </summary>
    public class DocxTableOfContents
    {
        private readonly List<TocEntry> _entries = new List<TocEntry>();
        private string _title = "Table of Contents";
        private bool _showNumbers = true;
        private bool _useLinkStyle = false;
        private string _titleAlignment = "left";
        private string _contentAlignment = "left";

        /// <summary>
        /// Sets the TOC title.
        /// </summary>
        public DocxTableOfContents SetTitle(string title)
        {
            _title = title;
            return this;
        }

        /// <summary>
        /// Whether to show serial numbers (1, 1.1, 1.2, 2, etc.). Default is true.
        /// </summary>
        public DocxTableOfContents ShowSerialNumbers(bool show = true)
        {
            _showNumbers = show;
            return this;
        }

        /// <summary>
        /// When true, TOC entries are displayed with blue underlined hyperlink style.
        /// When false (default), entries are displayed as normal black text.
        /// </summary>
        public DocxTableOfContents UseLinkStyle(bool useLinks = true)
        {
            _useLinkStyle = useLinks;
            return this;
        }

        /// <summary>
        /// Sets the alignment of the TOC title. Values: "left", "center", "right".
        /// </summary>
        public DocxTableOfContents SetTitleAlignment(string alignment)
        {
            _titleAlignment = alignment ?? "left";
            return this;
        }

        /// <summary>
        /// Sets the alignment of the TOC content entries. Values: "left", "center", "right".
        /// </summary>
        public DocxTableOfContents SetContentAlignment(string alignment)
        {
            _contentAlignment = alignment ?? "left";
            return this;
        }

        /// <summary>
        /// Adds a TOC entry.
        /// </summary>
        public DocxTableOfContents AddEntry(string title, int level, string bookmarkName)
        {
            _entries.Add(new TocEntry(title, level, bookmarkName));
            return this;
        }

        /// <summary>
        /// Gets the list of entries with computed serial numbers.
        /// Call this after adding all entries to get the serial number for each heading.
        /// </summary>
        public List<TocEntry> GetEntries()
        {
            ComputeSerialNumbers();
            return new List<TocEntry>(_entries);
        }

        private void ComputeSerialNumbers()
        {
            var counters = new int[10];
            foreach (var entry in _entries)
            {
                int level = entry.Level;
                if (level < 1) level = 1;
                if (level > 9) level = 9;

                counters[level]++;
                for (int i = level + 1; i < counters.Length; i++)
                    counters[i] = 0;

                if (_showNumbers)
                {
                    var parts = new List<string>();
                    for (int i = 1; i <= level; i++)
                        parts.Add(counters[i].ToString());
                    entry.SerialNumber = string.Join(".", parts);
                }
                else
                {
                    entry.SerialNumber = "";
                }
            }
        }

        internal string ToXml() => ToXml(null);

        internal string ToXml(Dictionary<string, int> bookmarkPages)
        {
            ComputeSerialNumbers();

            var sb = new StringBuilder();

            // TOC Title
            sb.Append("<w:p>");
            sb.Append("<w:pPr>");
            if (_titleAlignment != "left")
                sb.Append($"<w:jc w:val=\"{_titleAlignment}\"/>");
            sb.Append("</w:pPr>");
            sb.Append("<w:r><w:rPr><w:b/><w:sz w:val=\"32\"/></w:rPr>");
            sb.Append($"<w:t>{DocxDocument.EscapeXml(_title)}</w:t></w:r>");
            sb.Append("</w:p>");

            foreach (var entry in _entries)
            {
                int level = entry.Level;
                if (level < 1) level = 1;
                if (level > 9) level = 9;

                string displayText = _showNumbers && !string.IsNullOrEmpty(entry.SerialNumber)
                    ? entry.SerialNumber + "  " + entry.Title
                    : entry.Title;

                int indentTwips = (level - 1) * 360;

                sb.Append("<w:p>");
                sb.Append("<w:pPr>");
                sb.Append("<w:tabs>");
                sb.Append("<w:tab w:val=\"right\" w:leader=\"dot\" w:pos=\"9350\"/>");
                sb.Append("</w:tabs>");
                if (_contentAlignment != "left")
                    sb.Append($"<w:jc w:val=\"{_contentAlignment}\"/>");
                if (indentTwips > 0)
                    sb.Append($"<w:ind w:left=\"{indentTwips}\"/>");
                sb.Append("</w:pPr>");

                // Entry text with hyperlink to bookmark
                sb.Append($"<w:hyperlink w:anchor=\"{DocxDocument.EscapeXml(entry.BookmarkName)}\" w:history=\"1\">");
                sb.Append("<w:r>");

                // Style: blue link or normal black text
                if (_useLinkStyle)
                {
                    if (level == 1)
                        sb.Append("<w:rPr><w:rStyle w:val=\"Hyperlink\"/><w:b/></w:rPr>");
                    else
                        sb.Append("<w:rPr><w:rStyle w:val=\"Hyperlink\"/></w:rPr>");
                }
                else
                {
                    if (level == 1)
                        sb.Append("<w:rPr><w:b/><w:color w:val=\"000000\"/><w:u w:val=\"none\"/></w:rPr>");
                    else
                        sb.Append("<w:rPr><w:color w:val=\"000000\"/><w:u w:val=\"none\"/></w:rPr>");
                }

                sb.Append($"<w:t xml:space=\"preserve\">{DocxDocument.EscapeXml(displayText)}</w:t>");
                sb.Append("</w:r>");
                sb.Append("</w:hyperlink>");

                // Tab with dotted leader
                sb.Append("<w:r>");
                sb.Append("<w:tab/>");
                sb.Append("</w:r>");

                // PAGEREF field - cached value from tracked page number, w:dirty updates on open
                int cachedPage = 1;
                if (bookmarkPages != null && bookmarkPages.TryGetValue(entry.BookmarkName, out int tracked))
                    cachedPage = tracked;

                sb.Append("<w:r>");
                sb.Append("<w:fldChar w:fldCharType=\"begin\" w:dirty=\"true\"/>");
                sb.Append("</w:r>");
                sb.Append("<w:r>");
                sb.Append($"<w:instrText xml:space=\"preserve\"> PAGEREF {entry.BookmarkName} </w:instrText>");
                sb.Append("</w:r>");
                sb.Append("<w:r>");
                sb.Append("<w:fldChar w:fldCharType=\"separate\"/>");
                sb.Append("</w:r>");
                sb.Append("<w:r>");
                sb.Append("<w:rPr><w:noProof/></w:rPr>");
                sb.Append($"<w:t>{cachedPage}</w:t>");
                sb.Append("</w:r>");
                sb.Append("<w:r>");
                sb.Append("<w:fldChar w:fldCharType=\"end\"/>");
                sb.Append("</w:r>");

                sb.Append("</w:p>");
            }

            return sb.ToString();
        }
    }
}
