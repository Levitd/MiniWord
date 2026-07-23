using System;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Media.Imaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using MiniWord.Models;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using DxBold = DocumentFormat.OpenXml.Wordprocessing.Bold;
using DxColor = DocumentFormat.OpenXml.Wordprocessing.Color;
using DxItalic = DocumentFormat.OpenXml.Wordprocessing.Italic;
using DxPageSize = DocumentFormat.OpenXml.Wordprocessing.PageSize;
using DxPara = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using DxRun = DocumentFormat.OpenXml.Wordprocessing.Run;
using DxUnderline = DocumentFormat.OpenXml.Wordprocessing.Underline;
using MediaColor = System.Windows.Media.Color;
using MediaColors = System.Windows.Media.Colors;
using SolidBrush = System.Windows.Media.SolidColorBrush;
using WpfAlign = System.Windows.TextAlignment;
using WpfImage = System.Windows.Controls.Image;
using WpfList = System.Windows.Documents.List;
using WpfListItem = System.Windows.Documents.ListItem;
using WpfPara = System.Windows.Documents.Paragraph;
using WpfRun = System.Windows.Documents.Run;

namespace MiniWord.Services
{
    public class DocxService
    {
        // Unit conversions:
        //   WPF works in DIP (1/96 inch)
        //   docx font size: half-points; indents/spacing: twips (1/20 pt); images: EMU (914400/inch)
        private static double HalfPointsToDip(double halfPoints) => halfPoints * 2.0 / 3.0;
        private static int DipToHalfPoints(double dip) => (int)Math.Round(dip * 1.5);
        private static double TwipsToDip(double twips) => twips / 15.0;
        private static int DipToTwips(double dip) => (int)Math.Round(dip * 15.0);
        private static long DipToEmu(double dip) => (long)(dip / 96.0 * 914400.0);
        private static double EmuToDip(long emu) => emu / 914400.0 * 96.0;

        private const int BulletNumId = 1;
        private const int DecimalNumId = 2;

        /// <summary>Page size found in the last loaded document (null if none).</summary>
        public PageSizeInfo? LoadedPageSize { get; private set; }

        /// <summary>Header/footer settings found in the last loaded document.</summary>
        public DocumentSettings? LoadedSettings { get; private set; }

        #region Load

        public FlowDocument LoadDocument(string filePath)
        {
            LoadedPageSize = null;
            LoadedSettings = null;
            try
            {
                using var package = WordprocessingDocument.Open(filePath, false);
                var mainPart = package.MainDocumentPart;
                var root = mainPart?.Document;
                if (mainPart == null || root?.Body == null)
                    return new FlowDocument();

                var flowDoc = new FlowDocument();

                var pgSz = root.Body.Elements<SectionProperties>()
                    .Select(s => s.GetFirstChild<DxPageSize>())
                    .FirstOrDefault(p => p != null);
                if (pgSz?.Width != null)
                    LoadedPageSize = PageSizeInfo.ByTwips(pgSz.Width.Value);

                LoadedSettings = ReadHeaderFooter(mainPart);

                // Group consecutive numbered paragraphs into WPF lists
                WpfList? currentList = null;
                int currentNumId = -1;

                foreach (var element in root.Body.ChildElements)
                {
                    if (element is not DxPara dxPara)
                        continue;

                    var flowPara = BuildFlowParagraph(dxPara, mainPart);
                    var numPr = dxPara.ParagraphProperties?.NumberingProperties;

                    if (numPr?.NumberingId?.Val != null)
                    {
                        int numId = numPr.NumberingId.Val.Value;
                        if (currentList == null || numId != currentNumId)
                        {
                            currentList = new WpfList { MarkerStyle = ResolveMarker(mainPart, numId) };
                            currentNumId = numId;
                            flowDoc.Blocks.Add(currentList);
                        }
                        flowPara.Margin = new Thickness(0);
                        currentList.ListItems.Add(new WpfListItem(flowPara));
                    }
                    else
                    {
                        currentList = null;
                        flowDoc.Blocks.Add(flowPara);
                    }
                }

                if (flowDoc.Blocks.Count == 0)
                    flowDoc.Blocks.Add(new WpfPara());

                return flowDoc;
            }
            catch
            {
                return new FlowDocument(new WpfPara());
            }
        }

        private static WpfPara BuildFlowParagraph(DxPara dxPara, MainDocumentPart mainPart)
        {
            var flowPara = new WpfPara();
            ReadParagraphProperties(dxPara.ParagraphProperties, flowPara);

            foreach (var dxRun in dxPara.Elements<DxRun>())
            {
                foreach (var drawing in dxRun.Descendants<Drawing>())
                {
                    var image = ReadImage(mainPart, drawing);
                    if (image != null)
                        flowPara.Inlines.Add(new InlineUIContainer(image));
                }

                var text = string.Concat(dxRun.Elements<Text>().Select(t => t.Text));
                if (text.Length > 0)
                {
                    var flowRun = new WpfRun(text);
                    ReadRunProperties(dxRun.RunProperties, flowRun);
                    flowPara.Inlines.Add(flowRun);
                }
            }

            return flowPara;
        }

        private static TextMarkerStyle ResolveMarker(MainDocumentPart mainPart, int numId)
        {
            try
            {
                var numbering = mainPart.NumberingDefinitionsPart?.Numbering;
                if (numbering == null)
                    return TextMarkerStyle.Disc;

                var instance = numbering.Elements<NumberingInstance>()
                    .FirstOrDefault(n => n.NumberID?.Value == numId);
                var absId = instance?.GetFirstChild<AbstractNumId>()?.Val?.Value;
                var abs = numbering.Elements<AbstractNum>()
                    .FirstOrDefault(a => a.AbstractNumberId?.Value == absId);
                var level0 = abs?.Elements<Level>()
                    .FirstOrDefault(l => l.LevelIndex != null && l.LevelIndex.Value == 0);

                var fmt = level0?.NumberingFormat?.Val;
                if (fmt != null && fmt.Value == NumberFormatValues.Decimal)
                    return TextMarkerStyle.Decimal;
                return TextMarkerStyle.Disc;
            }
            catch
            {
                return TextMarkerStyle.Disc;
            }
        }

        private static DocumentSettings ReadHeaderFooter(MainDocumentPart mainPart)
        {
            var settings = new DocumentSettings();

            var header = mainPart.HeaderParts.FirstOrDefault()?.Header;
            if (header != null)
            {
                foreach (var para in header.Elements<DxPara>())
                {
                    if (HasPageField(para))
                    {
                        settings.ShowPageNumbers = true;
                        settings.PageNumberPosition = PageNumberPosition.HeaderRight;
                    }
                    else
                    {
                        var text = GetParagraphText(para);
                        if (text.Length > 0)
                            settings.HeaderText = settings.HeaderText.Length == 0 ? text : settings.HeaderText + " " + text;
                    }
                }
            }

            var footer = mainPart.FooterParts.FirstOrDefault()?.Footer;
            if (footer != null)
            {
                foreach (var para in footer.Elements<DxPara>())
                {
                    if (HasPageField(para))
                    {
                        settings.ShowPageNumbers = true;
                        var just = para.ParagraphProperties?.Justification?.Val;
                        settings.PageNumberPosition = just != null && just.Value == JustificationValues.Right
                            ? PageNumberPosition.FooterRight
                            : PageNumberPosition.FooterCenter;
                    }
                    else
                    {
                        var text = GetParagraphText(para);
                        if (text.Length > 0)
                            settings.FooterText = settings.FooterText.Length == 0 ? text : settings.FooterText + " " + text;
                    }
                }
            }

            return settings;
        }

        private static bool HasPageField(DxPara para) =>
            para.Descendants<SimpleField>().Any(f => f.Instruction?.Value?.Contains("PAGE") == true)
            || para.Descendants<FieldCode>().Any(f => f.Text.Contains("PAGE"));

        private static string GetParagraphText(DxPara para) =>
            string.Concat(para.Descendants<Text>().Select(t => t.Text)).Trim();

        private static void ReadParagraphProperties(ParagraphProperties? pPr, WpfPara flowPara)
        {
            if (pPr == null)
                return;

            if (pPr.Justification?.Val != null)
            {
                var justVal = pPr.Justification.Val.Value;
                if (justVal == JustificationValues.Center)
                    flowPara.TextAlignment = WpfAlign.Center;
                else if (justVal == JustificationValues.Right)
                    flowPara.TextAlignment = WpfAlign.Right;
                else if (justVal == JustificationValues.Both || justVal == JustificationValues.Distribute)
                    flowPara.TextAlignment = WpfAlign.Justify;
                else
                    flowPara.TextAlignment = WpfAlign.Left;
            }

            double left = 0, right = 0, before = 0, after = 0;

            var ind = pPr.Indentation;
            if (ind?.Left != null && double.TryParse(ind.Left.Value, out var l))
                left = TwipsToDip(l);
            if (ind?.Right != null && double.TryParse(ind.Right.Value, out var r))
                right = TwipsToDip(r);
            if (ind?.FirstLine != null && double.TryParse(ind.FirstLine.Value, out var fl))
                flowPara.TextIndent = TwipsToDip(fl);

            var spacing = pPr.SpacingBetweenLines;
            if (spacing?.Before != null && double.TryParse(spacing.Before.Value, out var b))
                before = TwipsToDip(b);
            if (spacing?.After != null && double.TryParse(spacing.After.Value, out var a))
                after = TwipsToDip(a);
            if (spacing?.Line != null && spacing.LineRule != null
                && spacing.LineRule.Value == LineSpacingRuleValues.Exact
                && double.TryParse(spacing.Line.Value, out var line))
                flowPara.LineHeight = TwipsToDip(line);

            flowPara.Margin = new Thickness(left, before, right, after);
        }

        private static void ReadRunProperties(RunProperties? rPr, WpfRun flowRun)
        {
            if (rPr == null)
                return;

            if (rPr.Bold != null)
                flowRun.FontWeight = FontWeights.Bold;
            if (rPr.Italic != null)
                flowRun.FontStyle = FontStyles.Italic;
            if (rPr.Underline != null)
                flowRun.TextDecorations = TextDecorations.Underline;

            if (rPr.FontSize?.Val != null && double.TryParse(rPr.FontSize.Val.Value, out var halfPoints))
                flowRun.FontSize = HalfPointsToDip(halfPoints);

            if (rPr.RunFonts?.Ascii != null)
                flowRun.FontFamily = new System.Windows.Media.FontFamily(rPr.RunFonts.Ascii);

            if (rPr.Color?.Val?.Value is string hex && TryParseHex(hex, out var fgColor))
                flowRun.Foreground = new SolidBrush(fgColor);

            if (rPr.Shading?.Fill?.Value is string fill && TryParseHex(fill, out var bgColor))
                flowRun.Background = new SolidBrush(bgColor);
            else if (rPr.Highlight?.Val != null)
            {
                var highlight = HighlightToColor(rPr.Highlight.Val.Value);
                if (highlight != null)
                    flowRun.Background = new SolidBrush(highlight.Value);
            }
        }

        private static bool TryParseHex(string hex, out MediaColor color)
        {
            color = MediaColors.Black;
            if (hex.Length != 6 || hex == "auto")
                return false;
            try
            {
                color = MediaColor.FromRgb(
                    Convert.ToByte(hex.Substring(0, 2), 16),
                    Convert.ToByte(hex.Substring(2, 2), 16),
                    Convert.ToByte(hex.Substring(4, 2), 16));
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static string ToHex(MediaColor c) => $"{c.R:X2}{c.G:X2}{c.B:X2}";

        private static MediaColor? HighlightToColor(HighlightColorValues v)
        {
            if (v == HighlightColorValues.Yellow) return MediaColors.Yellow;
            if (v == HighlightColorValues.Green) return MediaColor.FromRgb(0x00, 0xFF, 0x00);
            if (v == HighlightColorValues.Cyan) return MediaColors.Cyan;
            if (v == HighlightColorValues.Magenta) return MediaColors.Magenta;
            if (v == HighlightColorValues.Blue) return MediaColor.FromRgb(0x00, 0x00, 0xFF);
            if (v == HighlightColorValues.Red) return MediaColors.Red;
            if (v == HighlightColorValues.DarkBlue) return MediaColor.FromRgb(0x00, 0x00, 0x8B);
            if (v == HighlightColorValues.DarkCyan) return MediaColor.FromRgb(0x00, 0x8B, 0x8B);
            if (v == HighlightColorValues.DarkGreen) return MediaColor.FromRgb(0x00, 0x64, 0x00);
            if (v == HighlightColorValues.DarkMagenta) return MediaColor.FromRgb(0x8B, 0x00, 0x8B);
            if (v == HighlightColorValues.DarkRed) return MediaColor.FromRgb(0x8B, 0x00, 0x00);
            if (v == HighlightColorValues.DarkYellow) return MediaColor.FromRgb(0x80, 0x80, 0x00);
            if (v == HighlightColorValues.DarkGray) return MediaColor.FromRgb(0x80, 0x80, 0x80);
            if (v == HighlightColorValues.LightGray) return MediaColor.FromRgb(0xD3, 0xD3, 0xD3);
            if (v == HighlightColorValues.Black) return MediaColors.Black;
            if (v == HighlightColorValues.White) return MediaColors.White;
            return null;
        }

        private static WpfImage? ReadImage(MainDocumentPart mainPart, Drawing drawing)
        {
            try
            {
                var blip = drawing.Descendants<A.Blip>().FirstOrDefault();
                if (blip?.Embed?.Value == null)
                    return null;
                if (mainPart.GetPartById(blip.Embed.Value) is not ImagePart imagePart)
                    return null;

                var ms = new MemoryStream();
                using (var stream = imagePart.GetStream())
                    stream.CopyTo(ms);
                ms.Position = 0;

                var bmp = new BitmapImage();
                bmp.BeginInit();
                bmp.CacheOption = BitmapCacheOption.OnLoad;
                bmp.StreamSource = ms;
                bmp.EndInit();
                bmp.Freeze();

                double w = bmp.Width, h = bmp.Height;
                var extent = drawing.Descendants<DW.Extent>().FirstOrDefault();
                if (extent?.Cx != null && extent.Cy != null)
                {
                    w = EmuToDip(extent.Cx.Value);
                    h = EmuToDip(extent.Cy.Value);
                }

                return new WpfImage
                {
                    Source = bmp,
                    Width = w,
                    Height = h,
                    Stretch = System.Windows.Media.Stretch.Uniform
                };
            }
            catch
            {
                return null;
            }
        }

        #endregion

        #region Save

        public void SaveDocument(FlowDocument document, string filePath, PageSizeInfo pageSize, DocumentSettings? settings = null)
        {
            try
            {
                using var package = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document);
                var mainPart = package.AddMainDocumentPart();
                var body = new Body();
                uint imageId = 1;
                bool hasLists = document.Blocks.OfType<WpfList>().Any();

                if (hasLists)
                    AddNumberingPart(mainPart);

                foreach (var block in document.Blocks)
                    AppendBlock(body, block, mainPart, ref imageId, 0);

                var sectPr = new SectionProperties();
                if (settings != null && settings.HasAnyContent)
                    AddHeaderFooter(mainPart, settings, sectPr);
                sectPr.Append(new DxPageSize { Width = pageSize.WidthTwips, Height = pageSize.HeightTwips });
                sectPr.Append(new PageMargin
                {
                    Top = 1200,
                    Bottom = 1200,
                    Left = 1200U,
                    Right = 1200U,
                    Header = 720U,
                    Footer = 720U,
                    Gutter = 0U
                });
                body.Append(sectPr);

                mainPart.Document = new Document(body);
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to save document: {ex.Message}");
            }
        }

        private void AppendBlock(Body body, Block block, MainDocumentPart mainPart, ref uint imageId, int listLevel)
        {
            if (block is WpfPara flowPara)
            {
                body.Append(BuildDxParagraph(flowPara, mainPart, ref imageId, null, 0));
            }
            else if (block is WpfList list)
            {
                AppendList(body, list, mainPart, ref imageId, listLevel);
            }
            else if (block is BlockUIContainer container && container.Child is WpfImage image)
            {
                var dxPara = new DxPara();
                var imageRun = BuildImageRun(mainPart, image, imageId);
                if (imageRun != null)
                {
                    dxPara.Append(imageRun);
                    imageId++;
                }
                body.Append(dxPara);
            }
        }

        private void AppendList(Body body, WpfList list, MainDocumentPart mainPart, ref uint imageId, int level)
        {
            int numId = list.MarkerStyle == TextMarkerStyle.Decimal ? DecimalNumId : BulletNumId;
            int cappedLevel = Math.Min(level, 2);

            foreach (var item in list.ListItems)
            {
                foreach (var child in item.Blocks)
                {
                    if (child is WpfPara para)
                        body.Append(BuildDxParagraph(para, mainPart, ref imageId, numId, cappedLevel));
                    else if (child is WpfList nested)
                        AppendList(body, nested, mainPart, ref imageId, level + 1);
                }
            }
        }

        private DxPara BuildDxParagraph(WpfPara flowPara, MainDocumentPart mainPart, ref uint imageId, int? numId, int listLevel)
        {
            var dxPara = new DxPara();
            var pPr = BuildParagraphProperties(flowPara);

            if (numId != null)
            {
                pPr.InsertAt(new NumberingProperties(
                    new NumberingLevelReference { Val = listLevel },
                    new NumberingId { Val = numId.Value }), 0);
            }

            dxPara.Append(pPr);

            foreach (var inline in flowPara.Inlines)
            {
                if (inline is WpfRun flowRun)
                {
                    dxPara.Append(BuildTextRun(flowRun));
                }
                else if (inline is InlineUIContainer container && container.Child is WpfImage image)
                {
                    var imageRun = BuildImageRun(mainPart, image, imageId);
                    if (imageRun != null)
                    {
                        dxPara.Append(imageRun);
                        imageId++;
                    }
                }
            }

            return dxPara;
        }

        private static ParagraphProperties BuildParagraphProperties(WpfPara flowPara)
        {
            var pPr = new ParagraphProperties();

            JustificationValues justVal;
            if (flowPara.TextAlignment == WpfAlign.Center)
                justVal = JustificationValues.Center;
            else if (flowPara.TextAlignment == WpfAlign.Right)
                justVal = JustificationValues.Right;
            else if (flowPara.TextAlignment == WpfAlign.Justify)
                justVal = JustificationValues.Both;
            else
                justVal = JustificationValues.Left;
            pPr.Append(new Justification { Val = justVal });

            var m = flowPara.Margin;
            if (m.Left > 0 || m.Right > 0 || flowPara.TextIndent > 0)
            {
                var ind = new Indentation();
                if (m.Left > 0) ind.Left = DipToTwips(m.Left).ToString();
                if (m.Right > 0) ind.Right = DipToTwips(m.Right).ToString();
                if (flowPara.TextIndent > 0) ind.FirstLine = DipToTwips(flowPara.TextIndent).ToString();
                pPr.Append(ind);
            }

            var spacing = new SpacingBetweenLines();
            if (m.Top > 0) spacing.Before = DipToTwips(m.Top).ToString();
            if (m.Bottom > 0) spacing.After = DipToTwips(m.Bottom).ToString();
            if (flowPara.LineHeight > 0 && !double.IsNaN(flowPara.LineHeight))
            {
                spacing.Line = DipToTwips(flowPara.LineHeight).ToString();
                spacing.LineRule = LineSpacingRuleValues.Exact;
            }
            pPr.Append(spacing);

            return pPr;
        }

        private static DxRun BuildTextRun(WpfRun flowRun)
        {
            var dxRun = new DxRun();
            var rPr = new RunProperties();

            if (flowRun.FontFamily != null && !string.IsNullOrEmpty(flowRun.FontFamily.Source))
            {
                var font = flowRun.FontFamily.Source;
                rPr.Append(new RunFonts { Ascii = font, HighAnsi = font, ComplexScript = font });
            }

            if (flowRun.FontWeight == FontWeights.Bold)
                rPr.Append(new DxBold());

            if (flowRun.FontStyle == FontStyles.Italic)
                rPr.Append(new DxItalic());

            if (flowRun.TextDecorations?.Count > 0)
                rPr.Append(new DxUnderline { Val = UnderlineValues.Single });

            // Only save an explicitly set text color, not the inherited default
            if (flowRun.ReadLocalValue(TextElement.ForegroundProperty) != DependencyProperty.UnsetValue
                && flowRun.Foreground is SolidBrush fg && fg.Color.A == 255)
                rPr.Append(new DxColor { Val = ToHex(fg.Color) });

            if (flowRun.Background is SolidBrush bg && bg.Color.A == 255)
                rPr.Append(new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = ToHex(bg.Color) });

            if (flowRun.FontSize > 0)
                rPr.Append(new FontSize { Val = DipToHalfPoints(flowRun.FontSize).ToString() });

            dxRun.Append(rPr);
            dxRun.Append(new Text(flowRun.Text) { Space = SpaceProcessingModeValues.Preserve });
            return dxRun;
        }

        private static DxRun? BuildImageRun(MainDocumentPart mainPart, WpfImage image, uint id)
        {
            if (image.Source is not BitmapSource bmpSource)
                return null;

            var imagePart = mainPart.AddImagePart(ImagePartType.Png);
            using (var ms = new MemoryStream())
            {
                var encoder = new PngBitmapEncoder();
                encoder.Frames.Add(BitmapFrame.Create(bmpSource));
                encoder.Save(ms);
                ms.Position = 0;
                imagePart.FeedData(ms);
            }
            var relId = mainPart.GetIdOfPart(imagePart);

            double wDip = !double.IsNaN(image.Width) && image.Width > 0 ? image.Width : bmpSource.Width;
            double hDip = !double.IsNaN(image.Height) && image.Height > 0 ? image.Height : bmpSource.Height;
            long cx = DipToEmu(wDip);
            long cy = DipToEmu(hDip);

            var drawing = new Drawing(
                new DW.Inline(
                    new DW.Extent { Cx = cx, Cy = cy },
                    new DW.EffectExtent { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                    new DW.DocProperties { Id = id, Name = $"Picture {id}" },
                    new DW.NonVisualGraphicFrameDrawingProperties(
                        new A.GraphicFrameLocks { NoChangeAspect = true }),
                    new A.Graphic(
                        new A.GraphicData(
                            new PIC.Picture(
                                new PIC.NonVisualPictureProperties(
                                    new PIC.NonVisualDrawingProperties { Id = 0U, Name = $"image{id}.png" },
                                    new PIC.NonVisualPictureDrawingProperties()),
                                new PIC.BlipFill(
                                    new A.Blip { Embed = relId },
                                    new A.Stretch(new A.FillRectangle())),
                                new PIC.ShapeProperties(
                                    new A.Transform2D(
                                        new A.Offset { X = 0L, Y = 0L },
                                        new A.Extents { Cx = cx, Cy = cy }),
                                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })))
                        { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }))
                {
                    DistanceFromTop = 0U,
                    DistanceFromBottom = 0U,
                    DistanceFromLeft = 0U,
                    DistanceFromRight = 0U
                });

            return new DxRun(drawing);
        }

        private static void AddNumberingPart(MainDocumentPart mainPart)
        {
            var numberingPart = mainPart.AddNewPart<NumberingDefinitionsPart>();
            numberingPart.Numbering = new Numbering(
                BuildAbstractNum(0, bullet: true),
                BuildAbstractNum(1, bullet: false),
                new NumberingInstance(new AbstractNumId { Val = 0 }) { NumberID = BulletNumId },
                new NumberingInstance(new AbstractNumId { Val = 1 }) { NumberID = DecimalNumId });
        }

        private static AbstractNum BuildAbstractNum(int id, bool bullet)
        {
            var abstractNum = new AbstractNum { AbstractNumberId = id };
            string[] bulletChars = { "•", "○", "▪" };

            for (int i = 0; i < 3; i++)
            {
                var level = new Level { LevelIndex = i };
                if (!bullet)
                    level.Append(new StartNumberingValue { Val = 1 });
                level.Append(new NumberingFormat { Val = bullet ? NumberFormatValues.Bullet : NumberFormatValues.Decimal });
                level.Append(new LevelText { Val = bullet ? bulletChars[i] : $"%{i + 1}." });
                level.Append(new LevelJustification { Val = LevelJustificationValues.Left });
                level.Append(new PreviousParagraphProperties(
                    new Indentation { Left = (720 * (i + 1)).ToString(), Hanging = "360" }));
                if (bullet)
                    level.Append(new NumberingSymbolRunProperties(
                        new RunFonts { Ascii = "Segoe UI Symbol", HighAnsi = "Segoe UI Symbol" }));
                abstractNum.Append(level);
            }

            return abstractNum;
        }

        private static void AddHeaderFooter(MainDocumentPart mainPart, DocumentSettings settings, SectionProperties sectPr)
        {
            bool headerNumbers = settings.ShowPageNumbers && settings.PageNumberPosition == PageNumberPosition.HeaderRight;
            bool footerNumbers = settings.ShowPageNumbers && !headerNumbers;

            if (settings.HeaderText.Length > 0 || headerNumbers)
            {
                var part = mainPart.AddNewPart<HeaderPart>();
                var header = new Header();
                if (settings.HeaderText.Length > 0)
                    header.Append(new DxPara(new DxRun(new Text(settings.HeaderText))));
                if (headerNumbers)
                    header.Append(BuildPageNumberParagraph(JustificationValues.Right));
                part.Header = header;
                sectPr.Append(new HeaderReference { Type = HeaderFooterValues.Default, Id = mainPart.GetIdOfPart(part) });
            }

            if (settings.FooterText.Length > 0 || footerNumbers)
            {
                var part = mainPart.AddNewPart<FooterPart>();
                var footer = new Footer();
                if (settings.FooterText.Length > 0)
                    footer.Append(new DxPara(new DxRun(new Text(settings.FooterText))));
                if (footerNumbers)
                    footer.Append(BuildPageNumberParagraph(
                        settings.PageNumberPosition == PageNumberPosition.FooterRight
                            ? JustificationValues.Right
                            : JustificationValues.Center));
                part.Footer = footer;
                sectPr.Append(new FooterReference { Type = HeaderFooterValues.Default, Id = mainPart.GetIdOfPart(part) });
            }
        }

        private static DxPara BuildPageNumberParagraph(JustificationValues justification) =>
            new(
                new ParagraphProperties(new Justification { Val = justification }),
                new SimpleField(new DxRun(new Text("1"))) { Instruction = " PAGE " });

        #endregion
    }
}
