using System;
using System.IO;

namespace DocumentGenerator
{
    /// <summary>
    /// Creates the correct document builder based on the output file extension.
    /// Supported extensions: <c>.docx</c>, <c>.pdf</c>
    /// </summary>
    public static class DocumentFactory
    {
        /// <summary>
        /// Returns an <see cref="IDocument"/> instance appropriate for the given file path.
        /// <list type="bullet">
        ///   <item><c>.docx</c> → <see cref="DocxDocument"/></item>
        ///   <item><c>.pdf</c>  → <see cref="PdfDocument"/></item>
        /// </list>
        /// </summary>
        /// <param name="filePath">
        /// Destination file path. The extension determines which generator is returned.
        /// Call <see cref="IDocument.Save"/> with the same path when done.
        /// </param>
        /// <exception cref="NotSupportedException">
        /// Thrown when the file extension is not <c>.docx</c> or <c>.pdf</c>.
        /// </exception>
        public static IDocument Create(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                throw new ArgumentNullException(nameof(filePath));

            string ext = Path.GetExtension(filePath).ToLowerInvariant();

            return ext switch
            {
                ".docx" => new DocxDocument(),
                ".pdf"  => new PdfDocument(),
                _       => throw new NotSupportedException(
                                $"Unsupported file extension '{ext}'. " +
                                "Use '.docx' for Word documents or '.pdf' for PDF documents.")
            };
        }
    }
}
