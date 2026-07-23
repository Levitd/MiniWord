namespace MiniWord.Models
{
    public enum PageNumberPosition
    {
        FooterCenter,
        FooterRight,
        HeaderRight
    }

    /// <summary>Per-document settings: headers, footers, page numbering.</summary>
    public class DocumentSettings
    {
        public string HeaderText { get; set; } = "";
        public string FooterText { get; set; } = "";
        public bool ShowPageNumbers { get; set; }
        public PageNumberPosition PageNumberPosition { get; set; } = PageNumberPosition.FooterCenter;

        public bool HasAnyContent =>
            ShowPageNumbers || HeaderText.Length > 0 || FooterText.Length > 0;
    }
}
