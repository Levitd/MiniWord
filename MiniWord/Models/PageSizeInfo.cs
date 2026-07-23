using System.Linq;

namespace MiniWord.Models
{
    public class PageSizeInfo
    {
        public string Name { get; }
        public double WidthDip { get; }
        public double HeightDip { get; }

        // .docx stores page dimensions in twips (1/1440 inch); WPF uses DIP (1/96 inch)
        public uint WidthTwips => (uint)(WidthDip / 96.0 * 1440);
        public uint HeightTwips => (uint)(HeightDip / 96.0 * 1440);

        private PageSizeInfo(string name, double widthDip, double heightDip)
        {
            Name = name;
            WidthDip = widthDip;
            HeightDip = heightDip;
        }

        public static readonly PageSizeInfo[] All =
        {
            new("A4", 794, 1123),      // 210 x 297 mm
            new("Letter", 816, 1056),  // 8.5 x 11 in
            new("A5", 559, 794),       // 148 x 210 mm
        };

        public static PageSizeInfo ByName(string name) =>
            All.FirstOrDefault(p => p.Name == name) ?? All[0];

        public static PageSizeInfo ByTwips(uint widthTwips)
        {
            // Match with tolerance: files saved by Word use exact metric twips
            foreach (var p in All)
            {
                if (System.Math.Abs((long)p.WidthTwips - widthTwips) < 100)
                    return p;
            }
            return All[0];
        }
    }
}
