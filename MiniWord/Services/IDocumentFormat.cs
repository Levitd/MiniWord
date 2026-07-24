using System.Windows.Documents;
using MiniWord.Models;

namespace MiniWord.Services
{
    /// <summary>
    /// Common shape for a document reader/writer plugged in by file extension.
    /// DocxService already matches this; RTF/TXT/ODT services implement it too so
    /// MainWindow can dispatch on the file extension without special-casing docx.
    /// </summary>
    public interface IDocumentFormat
    {
        /// <summary>Page size found in the last loaded document (null if the format has none).</summary>
        PageSizeInfo? LoadedPageSize { get; }

        /// <summary>Header/footer settings found in the last loaded document (null if none).</summary>
        DocumentSettings? LoadedSettings { get; }

        FlowDocument LoadDocument(string filePath);

        void SaveDocument(FlowDocument document, string filePath, PageSizeInfo pageSize, DocumentSettings? settings = null);
    }
}
