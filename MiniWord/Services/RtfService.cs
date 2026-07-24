using System.IO;
using System.Windows;
using System.Windows.Documents;
using MiniWord.Models;

namespace MiniWord.Services
{
    /// <summary>
    /// Rich Text Format read/write via WPF's built-in TextRange serializer.
    /// RTF keeps character/paragraph formatting and images, and opens in
    /// WordPad, Word, LibreOffice, macOS TextEdit and Google Docs.
    /// </summary>
    public class RtfService : IDocumentFormat
    {
        // RTF carries no MiniWord page-size / header-footer metadata.
        public PageSizeInfo? LoadedPageSize => null;
        public DocumentSettings? LoadedSettings => null;

        public FlowDocument LoadDocument(string filePath)
        {
            var doc = new FlowDocument();
            using var fs = File.OpenRead(filePath);
            var range = new TextRange(doc.ContentStart, doc.ContentEnd);
            range.Load(fs, DataFormats.Rtf);
            return doc;
        }

        public void SaveDocument(FlowDocument document, string filePath, PageSizeInfo pageSize, DocumentSettings? settings = null)
        {
            using var fs = File.Create(filePath);
            var range = new TextRange(document.ContentStart, document.ContentEnd);
            range.Save(fs, DataFormats.Rtf);
        }
    }
}
