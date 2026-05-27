using System;
using System.IO;
using System.Text;

namespace DocumentGenerator
{
    /// <summary>
    /// Represents an image to be embedded in a DocxDocument.
    /// </summary>
    public class DocxImage
    {
        private byte[] _imageData;
        private readonly int _widthEmu;
        private readonly int _heightEmu;
        private readonly string _alignment;
        private readonly int _spacingBeforePt;
        private readonly int _spacingAfterPt;

        /// <summary>
        /// The relationship ID assigned when added to the document.
        /// </summary>
        internal string RelationshipId { get; set; }

        /// <summary>
        /// The file name for the image within the docx package.
        /// </summary>
        public string FileName { get; private set; }

        /// <summary>
        /// Creates an image from a file path.
        /// </summary>
        /// <param name="imagePath">Path to the image file.</param>
        /// <param name="widthEmu">Width in EMUs (English Metric Units). 1 inch = 914400 EMUs.</param>
        /// <param name="heightEmu">Height in EMUs. 1 inch = 914400 EMUs.</param>
        /// <param name="alignment">Paragraph alignment: left, center, right.</param>
        /// <param name="spacingBeforePt">Space before in points.</param>
        /// <param name="spacingAfterPt">Space after in points.</param>
        public DocxImage(string imagePath, int widthEmu = 5000000, int heightEmu = 3000000, string alignment = null, int spacingBeforePt = 0, int spacingAfterPt = 0)
        {
            _imageData = File.ReadAllBytes(imagePath);
            FileName = Path.GetFileName(imagePath);
            _widthEmu = widthEmu;
            _heightEmu = heightEmu;
            _alignment = alignment;
            _spacingBeforePt = spacingBeforePt;
            _spacingAfterPt = spacingAfterPt;
        }

        /// <summary>
        /// Creates an image from a byte array.
        /// </summary>
        public DocxImage(byte[] imageData, string fileName, int widthEmu = 5000000, int heightEmu = 3000000, string alignment = null, int spacingBeforePt = 0, int spacingAfterPt = 0)
        {
            _imageData = imageData;
            FileName = fileName;
            _widthEmu = widthEmu;
            _heightEmu = heightEmu;
            _alignment = alignment;
            _spacingBeforePt = spacingBeforePt;
            _spacingAfterPt = spacingAfterPt;
        }

        internal byte[] GetImageData()
        {
            return _imageData;
        }

        internal string ToXml()
        {
            string uniqueId = Math.Abs(Guid.NewGuid().GetHashCode()).ToString();

            var sb = new StringBuilder();
            sb.Append("<w:p>");

            // Paragraph properties for alignment and spacing
            bool hasPProps = !string.IsNullOrEmpty(_alignment) || _spacingBeforePt > 0 || _spacingAfterPt > 0;
            if (hasPProps)
            {
                sb.Append("<w:pPr>");
                if (!string.IsNullOrEmpty(_alignment))
                    sb.Append($"<w:jc w:val=\"{_alignment}\"/>");
                if (_spacingBeforePt > 0 || _spacingAfterPt > 0)
                {
                    sb.Append("<w:spacing");
                    if (_spacingBeforePt > 0) sb.Append($" w:before=\"{_spacingBeforePt * 20}\"");
                    if (_spacingAfterPt > 0) sb.Append($" w:after=\"{_spacingAfterPt * 20}\"");
                    sb.Append("/>");
                }
                sb.Append("</w:pPr>");
            }

            sb.Append("<w:r>");
            sb.Append("<w:drawing>");
            sb.Append("<wp:inline distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\">");
            sb.Append($"<wp:extent cx=\"{_widthEmu}\" cy=\"{_heightEmu}\"/>");
            sb.Append($"<wp:docPr id=\"{uniqueId}\" name=\"Picture\"/>");
            sb.Append("<a:graphic>");
            sb.Append("<a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">");
            sb.Append("<pic:pic>");
            sb.Append("<pic:nvPicPr>");
            sb.Append($"<pic:cNvPr id=\"0\" name=\"{DocxDocument.EscapeXml(FileName)}\"/>");
            sb.Append("<pic:cNvPicPr/>");
            sb.Append("</pic:nvPicPr>");
            sb.Append("<pic:blipFill>");
            sb.Append($"<a:blip r:embed=\"{RelationshipId}\"/>");
            sb.Append("<a:stretch><a:fillRect/></a:stretch>");
            sb.Append("</pic:blipFill>");
            sb.Append("<pic:spPr>");
            sb.Append("<a:xfrm>");
            sb.Append("<a:off x=\"0\" y=\"0\"/>");
            sb.Append($"<a:ext cx=\"{_widthEmu}\" cy=\"{_heightEmu}\"/>");
            sb.Append("</a:xfrm>");
            sb.Append("<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>");
            sb.Append("</pic:spPr>");
            sb.Append("</pic:pic>");
            sb.Append("</a:graphicData>");
            sb.Append("</a:graphic>");
            sb.Append("</wp:inline>");
            sb.Append("</w:drawing>");
            sb.Append("</w:r></w:p>");

            return sb.ToString();
        }
    }
}
