using System.Globalization;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Media;
using MiniWord.Models;

namespace MiniWord
{
    public class DocumentPaginator : System.Windows.Documents.DocumentPaginator
    {
        private const double PageMargin = 80;

        private readonly FlowDocument _document;
        private readonly DocumentSettings? _settings;
        private Size _pageSize;

        public DocumentPaginator(FlowDocument document, Size pageSize, DocumentSettings? settings = null)
        {
            _document = document;
            _pageSize = pageSize;
            _settings = settings;

            _document.PageWidth = pageSize.Width;
            _document.PageHeight = pageSize.Height;
            _document.PagePadding = new Thickness(PageMargin);
            _document.ColumnWidth = double.PositiveInfinity;

            var inner = Inner;
            inner.PageSize = pageSize;
            // Force synchronous pagination so PageCount is final before
            // the XPS writer starts asking for pages
            inner.ComputePageCount();
        }

        private System.Windows.Documents.DocumentPaginator Inner =>
            ((IDocumentPaginatorSource)_document).DocumentPaginator;

        public override DocumentPage GetPage(int pageNumber)
        {
            // Missing is the sentinel the XPS writer relies on to stop
            // enumerating pages — a blank page here means an infinite loop
            if (pageNumber >= Inner.PageCount)
                return DocumentPage.Missing;

            var page = Inner.GetPage(pageNumber);
            if (page == DocumentPage.Missing)
                return page;
            if (_settings == null || !_settings.HasAnyContent)
                return page;

            // Compose the original page with a header/footer overlay
            var container = new ContainerVisual();
            container.Children.Add(page.Visual);

            var overlay = new DrawingVisual();
            using (var dc = overlay.RenderOpen())
            {
                var typeface = new Typeface("Calibri");
                double footerY = _pageSize.Height - 55;

                if (_settings.HeaderText.Length > 0)
                    DrawText(dc, typeface, _settings.HeaderText, HPos.Left, 35);
                if (_settings.FooterText.Length > 0)
                    DrawText(dc, typeface, _settings.FooterText, HPos.Left, footerY);

                if (_settings.ShowPageNumbers)
                {
                    var num = (pageNumber + 1).ToString();
                    switch (_settings.PageNumberPosition)
                    {
                        case PageNumberPosition.HeaderRight:
                            DrawText(dc, typeface, num, HPos.Right, 35);
                            break;
                        case PageNumberPosition.FooterRight:
                            DrawText(dc, typeface, num, HPos.Right, footerY);
                            break;
                        default:
                            DrawText(dc, typeface, num, HPos.Center, footerY);
                            break;
                    }
                }
            }
            container.Children.Add(overlay);

            return new DocumentPage(container, page.Size, page.BleedBox, page.ContentBox);
        }

        private enum HPos { Left, Center, Right }

        private void DrawText(DrawingContext dc, Typeface typeface, string text, HPos pos, double y)
        {
            var ft = new FormattedText(text, CultureInfo.CurrentCulture, FlowDirection.LeftToRight,
                typeface, 13.3, Brushes.Gray, 1.0);
            double x = pos switch
            {
                HPos.Center => (_pageSize.Width - ft.Width) / 2,
                HPos.Right => _pageSize.Width - PageMargin - ft.Width,
                _ => PageMargin
            };
            dc.DrawText(ft, new Point(x, y));
        }

        public override bool IsPageCountValid => Inner.IsPageCountValid;

        public override int PageCount => Inner.PageCount;

        public override Size PageSize
        {
            get => _pageSize;
            set => _pageSize = value;
        }

        public override IDocumentPaginatorSource Source => _document;
    }
}
