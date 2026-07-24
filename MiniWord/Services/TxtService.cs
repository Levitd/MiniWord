using System.IO;
using System.Text;
using System.Windows.Documents;
using MiniWord.Models;

namespace MiniWord.Services
{
    /// <summary>
    /// Plain-text read/write. Formatting is intentionally dropped (that is what
    /// .txt means); every line becomes its own paragraph.
    /// </summary>
    public class TxtService : IDocumentFormat
    {
        public PageSizeInfo? LoadedPageSize => null;
        public DocumentSettings? LoadedSettings => null;

        public FlowDocument LoadDocument(string filePath)
        {
            var text = File.ReadAllText(filePath);
            var doc = new FlowDocument();
            foreach (var line in text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n'))
                doc.Blocks.Add(new Paragraph(new Run(line)));
            if (doc.Blocks.Count == 0)
                doc.Blocks.Add(new Paragraph());
            return doc;
        }

        public void SaveDocument(FlowDocument document, string filePath, PageSizeInfo pageSize, DocumentSettings? settings = null)
        {
            var range = new TextRange(document.ContentStart, document.ContentEnd);
            File.WriteAllText(filePath, range.Text, new UTF8Encoding(false));
        }
    }
}
