using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using MiniWord.Models;
using WpfPara = System.Windows.Documents.Paragraph;
using WpfRun = System.Windows.Documents.Run;
using WpfSpan = System.Windows.Documents.Span;
using WpfList = System.Windows.Documents.List;
using WpfHyperlink = System.Windows.Documents.Hyperlink;
using WpfLineBreak = System.Windows.Documents.LineBreak;
using WpfImage = System.Windows.Controls.Image;

namespace MiniWord.Services
{
    /// <summary>
    /// Export-only: renders the live FlowDocument to a single self-contained HTML
    /// file (inline CSS, images embedded as base64 data URIs) that opens in any
    /// browser. Importing HTML back is intentionally not supported.
    /// </summary>
    public static class HtmlExportService
    {
        static readonly CultureInfo Inv = CultureInfo.InvariantCulture;

        public static void Export(FlowDocument document, string filePath, PageSizeInfo pageSize)
        {
            var sb = new StringBuilder();
            sb.Append("<!DOCTYPE html>\n<html lang=\"en\">\n<head>\n<meta charset=\"utf-8\">\n");
            sb.Append("<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">\n");
            sb.Append("<title>Document</title>\n");
            sb.Append("<style>\n");
            sb.Append("  body{font-family:Calibri,'Segoe UI',sans-serif;font-size:16px;line-height:1.4;");
            sb.Append("color:#000;max-width:800px;margin:24px auto;padding:0 24px;}\n");
            sb.Append("  p{margin:0 0 8px 0;}\n  a{color:#0563C1;}\n  img{max-width:100%;height:auto;}\n");
            sb.Append("</style>\n</head>\n<body>\n");

            foreach (var block in document.Blocks)
                WriteBlock(sb, block);

            sb.Append("</body>\n</html>\n");
            File.WriteAllText(filePath, sb.ToString(), new UTF8Encoding(false));
        }

        private static void WriteBlock(StringBuilder sb, Block block)
        {
            if (block is WpfPara para)
            {
                string align = para.TextAlignment switch
                {
                    TextAlignment.Center => "center",
                    TextAlignment.Right => "right",
                    TextAlignment.Justify => "justify",
                    _ => ""
                };
                sb.Append(align.Length > 0 ? $"<p style=\"text-align:{align}\">" : "<p>");
                WriteInlines(sb, para.Inlines);
                sb.Append("</p>\n");
            }
            else if (block is WpfList list)
            {
                bool numbered = list.MarkerStyle == TextMarkerStyle.Decimal ||
                                list.MarkerStyle == TextMarkerStyle.LowerLatin ||
                                list.MarkerStyle == TextMarkerStyle.UpperLatin ||
                                list.MarkerStyle == TextMarkerStyle.LowerRoman ||
                                list.MarkerStyle == TextMarkerStyle.UpperRoman;
                string tag = numbered ? "ol" : "ul";
                sb.Append($"<{tag}>\n");
                foreach (var item in list.ListItems)
                {
                    sb.Append("<li>");
                    foreach (var b in item.Blocks)
                        if (b is WpfPara ip)
                            WriteInlines(sb, ip.Inlines);
                    sb.Append("</li>\n");
                }
                sb.Append($"</{tag}>\n");
            }
        }

        private static void WriteInlines(StringBuilder sb, InlineCollection inlines)
        {
            foreach (var inline in inlines)
            {
                switch (inline)
                {
                    case WpfHyperlink link:
                        var href = WebUtility.HtmlEncode(link.NavigateUri?.ToString() ?? "#");
                        sb.Append($"<a href=\"{href}\">");
                        WriteInlines(sb, link.Inlines);
                        sb.Append("</a>");
                        break;
                    case WpfRun run:
                        WriteRun(sb, run);
                        break;
                    case WpfLineBreak:
                        sb.Append("<br>");
                        break;
                    case InlineUIContainer uic when uic.Child is WpfImage img:
                        WriteImage(sb, img);
                        break;
                    case WpfSpan span:
                        WriteInlines(sb, span.Inlines);
                        break;
                }
            }
        }

        private static void WriteRun(StringBuilder sb, WpfRun run)
        {
            string text = WebUtility.HtmlEncode(run.Text ?? "").Replace("\t", "&#9;");
            var range = new TextRange(run.ContentStart, run.ContentEnd);
            var css = new StringBuilder();

            if (range.GetPropertyValue(TextElement.FontWeightProperty) is FontWeight fw && fw.ToOpenTypeWeight() >= 600)
                css.Append("font-weight:bold;");
            if (range.GetPropertyValue(TextElement.FontStyleProperty) is FontStyle fs && (fs == FontStyles.Italic || fs == FontStyles.Oblique))
                css.Append("font-style:italic;");

            bool underline = false, strike = false;
            if (range.GetPropertyValue(Inline.TextDecorationsProperty) is TextDecorationCollection tdc)
            {
                underline = tdc.Any(d => d.Location == TextDecorationLocation.Underline);
                strike = tdc.Any(d => d.Location == TextDecorationLocation.Strikethrough);
            }
            if (underline && strike) css.Append("text-decoration:underline line-through;");
            else if (underline) css.Append("text-decoration:underline;");
            else if (strike) css.Append("text-decoration:line-through;");

            if (range.GetPropertyValue(TextElement.ForegroundProperty) is SolidColorBrush b &&
                !(b.Color.R == 0 && b.Color.G == 0 && b.Color.B == 0))
                css.Append($"color:#{b.Color.R:x2}{b.Color.G:x2}{b.Color.B:x2};");
            if (range.GetPropertyValue(TextElement.FontFamilyProperty) is FontFamily ff)
                css.Append($"font-family:'{ff.Source}';");
            if (range.GetPropertyValue(TextElement.FontSizeProperty) is double sz)
                css.Append($"font-size:{(sz * 72.0 / 96.0).ToString("0.##", Inv)}pt;");

            int vert = 0;
            if (range.GetPropertyValue(Inline.BaselineAlignmentProperty) is BaselineAlignment ba)
                vert = ba == BaselineAlignment.Superscript ? 1 : ba == BaselineAlignment.Subscript ? -1 : 0;

            string open = vert == 1 ? "sup" : vert == -1 ? "sub" : "span";
            if (css.Length > 0)
                sb.Append($"<{open} style=\"{css}\">{text}</{open}>");
            else if (vert != 0)
                sb.Append($"<{open}>{text}</{open}>");
            else
                sb.Append(text);
        }

        private static void WriteImage(StringBuilder sb, WpfImage img)
        {
            try
            {
                if (img.Source is BitmapSource bmp)
                {
                    var encoder = new PngBitmapEncoder();
                    encoder.Frames.Add(BitmapFrame.Create(bmp));
                    using var ms = new MemoryStream();
                    encoder.Save(ms);
                    string b64 = Convert.ToBase64String(ms.ToArray());
                    string w = double.IsNaN(img.Width) ? "" : $" width=\"{(int)img.Width}\"";
                    sb.Append($"<img{w} src=\"data:image/png;base64,{b64}\">");
                }
            }
            catch { /* skip unrenderable images */ }
        }
    }
}
