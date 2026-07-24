using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Media;
using System.Xml;
using System.Xml.Linq;
using MiniWord.Models;
using WpfPara = System.Windows.Documents.Paragraph;
using WpfRun = System.Windows.Documents.Run;
using WpfSpan = System.Windows.Documents.Span;
using WpfList = System.Windows.Documents.List;
using WpfListItem = System.Windows.Documents.ListItem;
using WpfHyperlink = System.Windows.Documents.Hyperlink;
using WpfLineBreak = System.Windows.Documents.LineBreak;

namespace MiniWord.Services
{
    /// <summary>
    /// OpenDocument Text (.odt) read/write for LibreOffice / OpenOffice
    /// compatibility. Implemented directly over the ODF zip (content.xml +
    /// styles.xml + manifest) using System.IO.Compression + XDocument — no extra
    /// dependency. Covers paragraphs, character/paragraph formatting
    /// (bold/italic/underline/strikethrough/color/font/size/super-subscript,
    /// alignment), hyperlinks and single-level lists. Images are not carried
    /// through .odt (kept in .docx/.rtf); that is a documented limitation.
    /// </summary>
    public class OdtService : IDocumentFormat
    {
        static readonly XNamespace Office = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
        static readonly XNamespace Style = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
        static readonly XNamespace Text = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
        static readonly XNamespace Fo = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
        static readonly XNamespace Svg = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
        static readonly XNamespace Xlink = "http://www.w3.org/1999/xlink";
        static readonly XNamespace Manifest = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0";

        static readonly CultureInfo Inv = CultureInfo.InvariantCulture;

        public PageSizeInfo? LoadedPageSize { get; private set; }
        public DocumentSettings? LoadedSettings { get; private set; }

        static double PtToDip(double pt) => pt * 96.0 / 72.0;
        static double DipToPt(double dip) => dip * 72.0 / 96.0;

        /// <summary>Character formatting; nullable fields mean "not specified".</summary>
        private struct RunFmt
        {
            public bool? Bold, Italic, Underline, Strike;
            public Color? Color;
            public string? Font;
            public double? SizeDip;
            public int? Vert; // 1 super, -1 sub

            public bool IsEmpty =>
                Bold != true && Italic != true && Underline != true && Strike != true &&
                Color == null && Font == null && SizeDip == null && (Vert ?? 0) == 0;
        }

        private class StyleData
        {
            public TextAlignment? Align;
            public RunFmt Run;
        }

        // ==================== SAVE ====================

        public void SaveDocument(FlowDocument document, string filePath, PageSizeInfo pageSize, DocumentSettings? settings = null)
        {
            var fonts = new List<string>();
            var autoStyles = new XElement(Office + "automatic-styles");
            var bodyContent = new List<XElement>();
            var paraCache = new Dictionary<string, string>();
            var textCache = new Dictionary<string, string>();
            var listCache = new Dictionary<string, string>();
            int pC = 0, tC = 0, lC = 0;

            foreach (var block in document.Blocks)
                AppendBlock(block, bodyContent, autoStyles, paraCache, textCache, listCache, fonts, ref pC, ref tC, ref lC);

            var fontDecls = new XElement(Office + "font-face-decls",
                fonts.Distinct().Select(f => new XElement(Style + "font-face",
                    new XAttribute(Style + "name", f),
                    new XAttribute(Svg + "font-family", f))));

            var root = new XElement(Office + "document-content",
                new XAttribute(XNamespace.Xmlns + "office", Office.NamespaceName),
                new XAttribute(XNamespace.Xmlns + "style", Style.NamespaceName),
                new XAttribute(XNamespace.Xmlns + "text", Text.NamespaceName),
                new XAttribute(XNamespace.Xmlns + "fo", Fo.NamespaceName),
                new XAttribute(XNamespace.Xmlns + "svg", Svg.NamespaceName),
                new XAttribute(XNamespace.Xmlns + "xlink", Xlink.NamespaceName),
                new XAttribute(Office + "version", "1.2"),
                fontDecls,
                autoStyles,
                new XElement(Office + "body", new XElement(Office + "text", bodyContent)));

            var content = new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), root);

            if (File.Exists(filePath))
                File.Delete(filePath);

            using var fs = new FileStream(filePath, FileMode.Create);
            using var zip = new ZipArchive(fs, ZipArchiveMode.Create);
            // The mimetype entry must come first and be stored uncompressed.
            WriteRaw(zip, "mimetype", Encoding.ASCII.GetBytes("application/vnd.oasis.opendocument.text"), CompressionLevel.NoCompression);
            WriteXml(zip, "content.xml", content);
            WriteXml(zip, "styles.xml", BuildStyles());
            WriteXml(zip, "META-INF/manifest.xml", BuildManifest());
        }

        private void AppendBlock(Block block, List<XElement> outList, XElement autoStyles,
            Dictionary<string, string> paraCache, Dictionary<string, string> textCache, Dictionary<string, string> listCache,
            List<string> fonts, ref int pC, ref int tC, ref int lC)
        {
            if (block is WpfPara para)
            {
                outList.Add(BuildParagraph(para, autoStyles, paraCache, textCache, fonts, ref pC, ref tC));
            }
            else if (block is WpfList list)
            {
                string listStyle = GetListStyle(list.MarkerStyle, autoStyles, listCache, ref lC);
                var listEl = new XElement(Text + "list", new XAttribute(Text + "style-name", listStyle));
                foreach (var item in list.ListItems)
                {
                    var li = new XElement(Text + "list-item");
                    foreach (var b in item.Blocks)
                        if (b is WpfPara ip)
                            li.Add(BuildParagraph(ip, autoStyles, paraCache, textCache, fonts, ref pC, ref tC));
                    listEl.Add(li);
                }
                outList.Add(listEl);
            }
        }

        private XElement BuildParagraph(WpfPara para, XElement autoStyles,
            Dictionary<string, string> paraCache, Dictionary<string, string> textCache, List<string> fonts, ref int pC, ref int tC)
        {
            var pel = new XElement(Text + "p");
            string? ps = GetParaStyle(para.TextAlignment, autoStyles, paraCache, ref pC);
            if (ps != null)
                pel.SetAttributeValue(Text + "style-name", ps);
            AppendInlines(para.Inlines, pel, autoStyles, textCache, fonts, ref tC);
            return pel;
        }

        private void AppendInlines(InlineCollection inlines, XElement parent, XElement autoStyles,
            Dictionary<string, string> textCache, List<string> fonts, ref int tC)
        {
            foreach (var inline in inlines)
            {
                switch (inline)
                {
                    case WpfHyperlink link:
                        var a = new XElement(Text + "a",
                            new XAttribute(Xlink + "type", "simple"),
                            new XAttribute(Xlink + "href", link.NavigateUri?.ToString() ?? ""));
                        AppendInlines(link.Inlines, a, autoStyles, textCache, fonts, ref tC);
                        parent.Add(a);
                        break;
                    case WpfRun run:
                        AppendRun(run, parent, autoStyles, textCache, fonts, ref tC);
                        break;
                    case WpfLineBreak:
                        parent.Add(new XElement(Text + "line-break"));
                        break;
                    case WpfSpan span:
                        AppendInlines(span.Inlines, parent, autoStyles, textCache, fonts, ref tC);
                        break;
                    // InlineUIContainer (images) intentionally skipped for .odt
                }
            }
        }

        private void AppendRun(WpfRun run, XElement parent, XElement autoStyles,
            Dictionary<string, string> textCache, List<string> fonts, ref int tC)
        {
            var fmt = ReadRunFmt(run);
            string? styleName = GetTextStyle(fmt, autoStyles, textCache, fonts, ref tC);
            XElement target = parent;
            if (styleName != null)
            {
                var span = new XElement(Text + "span", new XAttribute(Text + "style-name", styleName));
                parent.Add(span);
                target = span;
            }
            AppendText(target, run.Text ?? "");
        }

        // Emit text preserving tabs and consecutive spaces (ODF collapses runs of
        // whitespace unless encoded with text:s / text:tab).
        private void AppendText(XElement parent, string text)
        {
            int i = 0;
            var buf = new StringBuilder();
            void Flush() { if (buf.Length > 0) { parent.Add(new XText(buf.ToString())); buf.Clear(); } }
            while (i < text.Length)
            {
                char c = text[i];
                if (c == '\t')
                {
                    Flush();
                    parent.Add(new XElement(Text + "tab"));
                    i++;
                }
                else if (c == ' ')
                {
                    int n = 0;
                    while (i < text.Length && text[i] == ' ') { n++; i++; }
                    buf.Append(' '); // first space stays literal
                    if (n > 1)
                    {
                        Flush();
                        var s = new XElement(Text + "s");
                        if (n - 1 > 1) s.SetAttributeValue(Text + "c", n - 1);
                        parent.Add(s);
                    }
                }
                else
                {
                    buf.Append(c);
                    i++;
                }
            }
            Flush();
        }

        private RunFmt ReadRunFmt(WpfRun run)
        {
            var range = new TextRange(run.ContentStart, run.ContentEnd);
            var f = new RunFmt();
            if (range.GetPropertyValue(TextElement.FontWeightProperty) is FontWeight fw)
                f.Bold = fw.ToOpenTypeWeight() >= 600;
            if (range.GetPropertyValue(TextElement.FontStyleProperty) is FontStyle fs)
                f.Italic = fs == FontStyles.Italic || fs == FontStyles.Oblique;
            if (range.GetPropertyValue(Inline.TextDecorationsProperty) is TextDecorationCollection tdc)
            {
                f.Underline = tdc.Any(d => d.Location == TextDecorationLocation.Underline);
                f.Strike = tdc.Any(d => d.Location == TextDecorationLocation.Strikethrough);
            }
            if (range.GetPropertyValue(TextElement.ForegroundProperty) is SolidColorBrush b)
                f.Color = b.Color;
            if (range.GetPropertyValue(TextElement.FontFamilyProperty) is FontFamily ff)
                f.Font = ff.Source;
            if (range.GetPropertyValue(TextElement.FontSizeProperty) is double sz)
                f.SizeDip = sz;
            if (range.GetPropertyValue(Inline.BaselineAlignmentProperty) is BaselineAlignment ba)
                f.Vert = ba == BaselineAlignment.Superscript ? 1 : ba == BaselineAlignment.Subscript ? -1 : 0;
            return f;
        }

        private string? GetTextStyle(RunFmt f, XElement autoStyles, Dictionary<string, string> cache, List<string> fonts, ref int tC)
        {
            // Always carry font + size for fidelity; other props only when set.
            bool hasColor = f.Color.HasValue && !(f.Color.Value.R == 0 && f.Color.Value.G == 0 && f.Color.Value.B == 0);
            string key = string.Join("|",
                f.Bold == true, f.Italic == true, f.Underline == true, f.Strike == true, f.Vert ?? 0,
                hasColor ? ToHex(f.Color!.Value) : "", f.Font ?? "", f.SizeDip?.ToString("0.##", Inv) ?? "");

            if (cache.TryGetValue(key, out var existing))
                return existing;

            var tp = new XElement(Style + "text-properties");
            if (f.Bold == true) tp.SetAttributeValue(Fo + "font-weight", "bold");
            if (f.Italic == true) { tp.SetAttributeValue(Fo + "font-style", "italic"); }
            if (f.Underline == true)
            {
                tp.SetAttributeValue(Style + "text-underline-style", "solid");
                tp.SetAttributeValue(Style + "text-underline-width", "auto");
                tp.SetAttributeValue(Style + "text-underline-color", "font-color");
            }
            if (f.Strike == true)
                tp.SetAttributeValue(Style + "text-line-through-style", "solid");
            if ((f.Vert ?? 0) == 1) tp.SetAttributeValue(Style + "text-position", "super 58%");
            if ((f.Vert ?? 0) == -1) tp.SetAttributeValue(Style + "text-position", "sub 58%");
            if (hasColor) tp.SetAttributeValue(Fo + "color", ToHex(f.Color!.Value));
            if (!string.IsNullOrEmpty(f.Font))
            {
                tp.SetAttributeValue(Style + "font-name", f.Font);
                fonts.Add(f.Font);
            }
            if (f.SizeDip.HasValue)
                tp.SetAttributeValue(Fo + "font-size", DipToPt(f.SizeDip.Value).ToString("0.##", Inv) + "pt");

            if (!tp.HasAttributes)
                return null;

            string name = "T" + (++tC);
            autoStyles.Add(new XElement(Style + "style",
                new XAttribute(Style + "name", name),
                new XAttribute(Style + "family", "text"),
                tp));
            cache[key] = name;
            return name;
        }

        private string? GetParaStyle(TextAlignment align, XElement autoStyles, Dictionary<string, string> cache, ref int pC)
        {
            string? foAlign = align switch
            {
                TextAlignment.Center => "center",
                TextAlignment.Right => "end",
                TextAlignment.Justify => "justify",
                _ => null
            };
            if (foAlign == null)
                return null;

            string key = "align:" + foAlign;
            if (cache.TryGetValue(key, out var existing))
                return existing;

            string name = "P" + (++pC);
            autoStyles.Add(new XElement(Style + "style",
                new XAttribute(Style + "name", name),
                new XAttribute(Style + "family", "paragraph"),
                new XElement(Style + "paragraph-properties", new XAttribute(Fo + "text-align", foAlign))));
            cache[key] = name;
            return name;
        }

        private string GetListStyle(TextMarkerStyle marker, XElement autoStyles, Dictionary<string, string> cache, ref int lC)
        {
            bool numbered = marker == TextMarkerStyle.Decimal || marker == TextMarkerStyle.LowerLatin ||
                            marker == TextMarkerStyle.UpperLatin || marker == TextMarkerStyle.LowerRoman ||
                            marker == TextMarkerStyle.UpperRoman;
            string key = numbered ? "num" : "bullet";
            if (cache.TryGetValue(key, out var existing))
                return existing;

            string name = "L" + (++lC);
            XElement level = numbered
                ? new XElement(Text + "list-level-style-number",
                    new XAttribute(Text + "level", "1"),
                    new XAttribute(Style + "num-format", "1"),
                    new XAttribute(Style + "num-suffix", "."))
                : new XElement(Text + "list-level-style-bullet",
                    new XAttribute(Text + "level", "1"),
                    new XAttribute(Text + "bullet-char", "•"));
            autoStyles.Add(new XElement(Text + "list-style",
                new XAttribute(Style + "name", name),
                level));
            cache[key] = name;
            return name;
        }

        private static XDocument BuildStyles()
        {
            var root = new XElement(Office + "document-styles",
                new XAttribute(XNamespace.Xmlns + "office", Office.NamespaceName),
                new XAttribute(XNamespace.Xmlns + "style", Style.NamespaceName),
                new XAttribute(XNamespace.Xmlns + "fo", Fo.NamespaceName),
                new XAttribute(Office + "version", "1.2"),
                new XElement(Office + "styles",
                    new XElement(Style + "default-style",
                        new XAttribute(Style + "family", "paragraph"),
                        new XElement(Style + "text-properties",
                            new XAttribute(Fo + "font-size", "11pt"))),
                    new XElement(Style + "style",
                        new XAttribute(Style + "name", "Standard"),
                        new XAttribute(Style + "family", "paragraph"),
                        new XAttribute(Style + "class", "text"))));
            return new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), root);
        }

        private static XDocument BuildManifest()
        {
            var root = new XElement(Manifest + "manifest",
                new XAttribute(XNamespace.Xmlns + "manifest", Manifest.NamespaceName),
                new XAttribute(Manifest + "version", "1.2"),
                new XElement(Manifest + "file-entry",
                    new XAttribute(Manifest + "full-path", "/"),
                    new XAttribute(Manifest + "version", "1.2"),
                    new XAttribute(Manifest + "media-type", "application/vnd.oasis.opendocument.text")),
                new XElement(Manifest + "file-entry",
                    new XAttribute(Manifest + "full-path", "content.xml"),
                    new XAttribute(Manifest + "media-type", "text/xml")),
                new XElement(Manifest + "file-entry",
                    new XAttribute(Manifest + "full-path", "styles.xml"),
                    new XAttribute(Manifest + "media-type", "text/xml")));
            return new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), root);
        }

        private static void WriteRaw(ZipArchive zip, string name, byte[] data, CompressionLevel level)
        {
            var entry = zip.CreateEntry(name, level);
            using var s = entry.Open();
            s.Write(data, 0, data.Length);
        }

        private static void WriteXml(ZipArchive zip, string name, XDocument xml)
        {
            using var ms = new MemoryStream();
            var settings = new XmlWriterSettings { Encoding = new UTF8Encoding(false), Indent = false };
            using (var w = XmlWriter.Create(ms, settings))
                xml.Save(w);
            WriteRaw(zip, name, ms.ToArray(), CompressionLevel.Optimal);
        }

        private static string ToHex(Color c) => $"#{c.R:x2}{c.G:x2}{c.B:x2}";

        // ==================== LOAD ====================

        public FlowDocument LoadDocument(string filePath)
        {
            LoadedPageSize = null;
            LoadedSettings = null;
            var doc = new FlowDocument();
            try
            {
                using var fs = File.OpenRead(filePath);
                using var zip = new ZipArchive(fs, ZipArchiveMode.Read);

                var styles = new Dictionary<string, StyleData>();
                ParseStylesFrom(zip, "styles.xml", styles);
                ParseStylesFrom(zip, "content.xml", styles);

                var contentEntry = zip.GetEntry("content.xml");
                if (contentEntry == null)
                {
                    doc.Blocks.Add(new WpfPara());
                    return doc;
                }

                XDocument xdoc;
                using (var s = contentEntry.Open())
                    xdoc = XDocument.Load(s);

                var textRoot = xdoc.Root?.Element(Office + "body")?.Element(Office + "text");
                if (textRoot != null)
                {
                    foreach (var el in textRoot.Elements())
                        ReadBodyElement(el, doc.Blocks, styles);
                }

                if (doc.Blocks.Count == 0)
                    doc.Blocks.Add(new WpfPara());
                return doc;
            }
            catch
            {
                return new FlowDocument(new WpfPara());
            }
        }

        private void ParseStylesFrom(ZipArchive zip, string entryName, Dictionary<string, StyleData> styles)
        {
            var entry = zip.GetEntry(entryName);
            if (entry == null) return;
            XDocument xdoc;
            try { using var s = entry.Open(); xdoc = XDocument.Load(s); }
            catch { return; }

            foreach (var container in xdoc.Descendants(Office + "automatic-styles")
                         .Concat(xdoc.Descendants(Office + "styles")))
            {
                foreach (var st in container.Elements(Style + "style"))
                {
                    var name = (string?)st.Attribute(Style + "name");
                    if (string.IsNullOrEmpty(name)) continue;
                    var data = new StyleData();

                    var pp = st.Element(Style + "paragraph-properties");
                    var alignAttr = (string?)pp?.Attribute(Fo + "text-align");
                    data.Align = alignAttr switch
                    {
                        "center" => TextAlignment.Center,
                        "end" or "right" => TextAlignment.Right,
                        "justify" => TextAlignment.Justify,
                        "start" or "left" => TextAlignment.Left,
                        _ => null
                    };

                    data.Run = ReadTextProps(st.Element(Style + "text-properties"));
                    styles[name] = data;
                }
            }
        }

        private RunFmt ReadTextProps(XElement? tp)
        {
            var f = new RunFmt();
            if (tp == null) return f;

            var weight = (string?)tp.Attribute(Fo + "font-weight");
            if (weight != null) f.Bold = weight == "bold" || (int.TryParse(weight, out var w) && w >= 600);

            var italic = (string?)tp.Attribute(Fo + "font-style");
            if (italic != null) f.Italic = italic == "italic" || italic == "oblique";

            var underline = (string?)tp.Attribute(Style + "text-underline-style");
            if (underline != null) f.Underline = underline != "none";

            var strike = (string?)tp.Attribute(Style + "text-line-through-style");
            if (strike != null) f.Strike = strike != "none";

            var color = (string?)tp.Attribute(Fo + "color");
            if (color != null && TryParseColor(color, out var col)) f.Color = col;

            var font = (string?)tp.Attribute(Style + "font-name") ?? (string?)tp.Attribute(Fo + "font-family");
            if (!string.IsNullOrEmpty(font)) f.Font = font.Trim('\'', '"');

            var size = (string?)tp.Attribute(Fo + "font-size");
            if (size != null && size.EndsWith("pt") &&
                double.TryParse(size[..^2], NumberStyles.Float, Inv, out var pt))
                f.SizeDip = PtToDip(pt);

            var pos = (string?)tp.Attribute(Style + "text-position");
            if (pos != null)
            {
                if (pos.StartsWith("super", StringComparison.OrdinalIgnoreCase)) f.Vert = 1;
                else if (pos.StartsWith("sub", StringComparison.OrdinalIgnoreCase)) f.Vert = -1;
                else if (pos.StartsWith("-")) f.Vert = -1;
                else if (double.TryParse(pos.Split(' ')[0].TrimEnd('%'), NumberStyles.Float, Inv, out var pv))
                    f.Vert = pv > 0 ? 1 : pv < 0 ? -1 : 0;
            }
            return f;
        }

        private void ReadBodyElement(XElement el, BlockCollection blocks, Dictionary<string, StyleData> styles)
        {
            if (el.Name == Text + "p" || el.Name == Text + "h")
            {
                blocks.Add(ReadParagraph(el, styles));
            }
            else if (el.Name == Text + "list")
            {
                var wlist = new WpfList { MarkerStyle = DetectListMarker(el, styles) };
                foreach (var item in el.Elements(Text + "list-item"))
                {
                    var li = new WpfListItem();
                    foreach (var p in item.Elements())
                        if (p.Name == Text + "p" || p.Name == Text + "h")
                        {
                            var para = ReadParagraph(p, styles);
                            para.Margin = new Thickness(0);
                            li.Blocks.Add(para);
                        }
                    if (li.Blocks.Count == 0)
                        li.Blocks.Add(new WpfPara());
                    wlist.ListItems.Add(li);
                }
                if (wlist.ListItems.Count > 0)
                    blocks.Add(wlist);
            }
        }

        private TextMarkerStyle DetectListMarker(XElement listEl, Dictionary<string, StyleData> styles)
        {
            // Simplest reliable signal available on read: numbered lists reference a
            // list style whose name we recorded, but we didn't store list styles, so
            // fall back to bullet unless the list style name hints otherwise.
            var styleName = (string?)listEl.Attribute(Text + "style-name") ?? "";
            return styleName.Contains("num", StringComparison.OrdinalIgnoreCase)
                ? TextMarkerStyle.Decimal
                : TextMarkerStyle.Disc;
        }

        private WpfPara ReadParagraph(XElement el, Dictionary<string, StyleData> styles)
        {
            var para = new WpfPara();
            RunFmt baseFmt = default;
            var styleName = (string?)el.Attribute(Text + "style-name");
            if (styleName != null && styles.TryGetValue(styleName, out var sd))
            {
                if (sd.Align.HasValue) para.TextAlignment = sd.Align.Value;
                baseFmt = sd.Run;
            }
            ReadInlines(el, para.Inlines, styles, baseFmt);
            return para;
        }

        private void ReadInlines(XElement parent, InlineCollection inlines, Dictionary<string, StyleData> styles, RunFmt baseFmt)
        {
            foreach (var node in parent.Nodes())
            {
                if (node is XText t)
                {
                    inlines.Add(MakeRun(t.Value, baseFmt));
                }
                else if (node is XElement e)
                {
                    if (e.Name == Text + "span")
                    {
                        var fmt = baseFmt;
                        var sn = (string?)e.Attribute(Text + "style-name");
                        if (sn != null && styles.TryGetValue(sn, out var sd))
                            fmt = Merge(baseFmt, sd.Run);
                        ReadInlines(e, inlines, styles, fmt);
                    }
                    else if (e.Name == Text + "a")
                    {
                        var link = new WpfHyperlink();
                        var href = (string?)e.Attribute(Xlink + "href");
                        if (!string.IsNullOrEmpty(href) && Uri.TryCreate(href, UriKind.RelativeOrAbsolute, out var uri))
                            link.NavigateUri = uri;
                        ReadInlines(e, link.Inlines, styles, baseFmt);
                        if (link.Inlines.Count == 0)
                            link.Inlines.Add(new WpfRun(href ?? ""));
                        inlines.Add(link);
                    }
                    else if (e.Name == Text + "s")
                    {
                        int c = 1;
                        var ca = (string?)e.Attribute(Text + "c");
                        if (ca != null) int.TryParse(ca, out c);
                        inlines.Add(MakeRun(new string(' ', Math.Max(1, c)), baseFmt));
                    }
                    else if (e.Name == Text + "tab")
                    {
                        inlines.Add(MakeRun("\t", baseFmt));
                    }
                    else if (e.Name == Text + "line-break")
                    {
                        inlines.Add(new WpfLineBreak());
                    }
                }
            }
        }

        private static RunFmt Merge(RunFmt b, RunFmt o) => new RunFmt
        {
            Bold = o.Bold ?? b.Bold,
            Italic = o.Italic ?? b.Italic,
            Underline = o.Underline ?? b.Underline,
            Strike = o.Strike ?? b.Strike,
            Color = o.Color ?? b.Color,
            Font = o.Font ?? b.Font,
            SizeDip = o.SizeDip ?? b.SizeDip,
            Vert = o.Vert ?? b.Vert
        };

        private WpfRun MakeRun(string text, RunFmt f)
        {
            var run = new WpfRun(text);
            if (f.Bold == true) run.FontWeight = FontWeights.Bold;
            if (f.Italic == true) run.FontStyle = FontStyles.Italic;
            if (f.Underline == true || f.Strike == true)
            {
                var tdc = new TextDecorationCollection();
                if (f.Underline == true) tdc.Add(TextDecorations.Underline);
                if (f.Strike == true) tdc.Add(TextDecorations.Strikethrough);
                run.TextDecorations = tdc;
            }
            if (f.Color.HasValue) run.Foreground = new SolidColorBrush(f.Color.Value);
            if (!string.IsNullOrEmpty(f.Font)) run.FontFamily = new FontFamily(f.Font);
            if (f.SizeDip.HasValue) run.FontSize = f.SizeDip.Value;
            if ((f.Vert ?? 0) == 1) run.BaselineAlignment = BaselineAlignment.Superscript;
            else if ((f.Vert ?? 0) == -1) run.BaselineAlignment = BaselineAlignment.Subscript;
            return run;
        }

        private static bool TryParseColor(string hex, out Color color)
        {
            color = Colors.Black;
            hex = hex.Trim();
            if (hex.StartsWith("#")) hex = hex[1..];
            if (hex.Length == 6 &&
                byte.TryParse(hex[..2], NumberStyles.HexNumber, Inv, out var r) &&
                byte.TryParse(hex.Substring(2, 2), NumberStyles.HexNumber, Inv, out var g) &&
                byte.TryParse(hex.Substring(4, 2), NumberStyles.HexNumber, Inv, out var b))
            {
                color = Color.FromRgb(r, g, b);
                return true;
            }
            return false;
        }
    }
}
