using System.Windows;
using System.Windows.Controls;
using MiniWord.Services;

namespace MiniWord
{
    public partial class SymbolDialog : Window
    {
        public string? SelectedSymbol { get; private set; }

        // Common symbols people actually reach for
        private static readonly string[] Symbols =
        {
            "©", "®", "™", "§", "¶", "†", "‡", "•", "…", "‰",
            "°", "№", "±", "×", "÷", "≠", "≈", "≤", "≥", "∞",
            "€", "£", "¥", "¢", "$", "₽", "½", "¼", "¾", "⅓",
            "–", "—", "«", "»", "„", "“", "”", "‘", "’", "→",
            "←", "↑", "↓", "↔", "√", "∑", "∏", "∆", "π", "µ",
            "α", "β", "γ", "δ", "λ", "Ω", "★", "☆", "✓", "✗",
        };

        public SymbolDialog()
        {
            InitializeComponent();
            Title = Loc.T("SymbolTitle");
            CancelButton.Content = Loc.T("Cancel");

            foreach (var sym in Symbols)
            {
                var btn = new Button
                {
                    Content = sym,
                    FontSize = 18,
                    Width = 34,
                    Height = 34,
                    Margin = new Thickness(1)
                };
                btn.Click += (s, e) =>
                {
                    SelectedSymbol = sym;
                    DialogResult = true;
                    Close();
                };
                SymbolGrid.Children.Add(btn);
            }
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
