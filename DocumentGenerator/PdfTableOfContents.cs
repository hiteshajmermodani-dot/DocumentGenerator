using System.Collections.Generic;

namespace DocumentGenerator
{
    /// <summary>
    /// Builds a Table of Contents for inclusion in a PdfDocument.
    /// Mirrors DocxTableOfContents API.
    /// </summary>
    public class PdfTableOfContents
    {
        private readonly List<TocEntry> _entries = new List<TocEntry>();
        private bool _showNumbers = true;
        private bool _useLinkStyle = false;

        public string Title { get; private set; } = "Table of Contents";
        internal bool ShowNumbers => _showNumbers;
        internal bool LinkStyle => _useLinkStyle;
        internal List<TocEntry> Entries => _entries;

        public PdfTableOfContents SetTitle(string title)
        {
            Title = title;
            return this;
        }

        public PdfTableOfContents ShowSerialNumbers(bool show = true)
        {
            _showNumbers = show;
            return this;
        }

        public PdfTableOfContents UseLinkStyle(bool useLinks = true)
        {
            _useLinkStyle = useLinks;
            return this;
        }

        public PdfTableOfContents SetTitleAlignment(string alignment) => this;
        public PdfTableOfContents SetContentAlignment(string alignment) => this;

        public PdfTableOfContents AddEntry(string title, int level, string bookmarkName)
        {
            _entries.Add(new TocEntry(title, level, bookmarkName));
            return this;
        }

        public List<TocEntry> GetEntries()
        {
            ComputeSerialNumbers();
            return new List<TocEntry>(_entries);
        }

        internal void ComputeSerialNumbers()
        {
            var counters = new int[10];
            foreach (var entry in _entries)
            {
                int level = entry.Level < 1 ? 1 : entry.Level > 9 ? 9 : entry.Level;
                counters[level]++;
                for (int i = level + 1; i < counters.Length; i++) counters[i] = 0;

                if (_showNumbers)
                {
                    var parts = new List<string>();
                    for (int i = 1; i <= level; i++) parts.Add(counters[i].ToString());
                    entry.SerialNumber = string.Join(".", parts);
                }
                else
                {
                    entry.SerialNumber = "";
                }
            }
        }
    }
}
