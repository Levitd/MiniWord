using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using MiniWord.Services;

namespace MiniWord
{
    public partial class FindReplaceDialog : Window
    {
        private readonly RichTextBox _editor;

        public FindReplaceDialog(RichTextBox editor)
        {
            InitializeComponent();
            _editor = editor;

            Title = Loc.T("FindTitle");
            FindLabel.Content = Loc.T("FindWhat");
            ReplaceLabel.Content = Loc.T("ReplaceWith");
            MatchCaseCheck.Content = Loc.T("MatchCase");
            FindNextButton.Content = Loc.T("FindNext");
            ReplaceButton.Content = Loc.T("ReplaceBtn");
            ReplaceAllButton.Content = Loc.T("ReplaceAll");
            CloseButton.Content = Loc.T("Close");
        }

        public void SetReplaceVisible(bool visible)
        {
            ReplaceRow.Visibility = visible ? Visibility.Visible : Visibility.Collapsed;
            ReplaceButton.Visibility = visible ? Visibility.Visible : Visibility.Collapsed;
            ReplaceAllButton.Visibility = visible ? Visibility.Visible : Visibility.Collapsed;
            Height = visible ? 200 : 165;
            FindBox.Focus();
        }

        private StringComparison Comparison => MatchCaseCheck.IsChecked == true
            ? StringComparison.CurrentCulture
            : StringComparison.CurrentCultureIgnoreCase;

        // Finds the next match after the given pointer; wraps once from the top.
        private TextRange? FindFrom(TextPointer start, string query)
        {
            var match = SearchForward(start, query);
            if (match == null)
                match = SearchForward(_editor.Document.ContentStart, query);
            return match;
        }

        private TextRange? SearchForward(TextPointer from, string query)
        {
            for (TextPointer? p = from; p != null; p = p.GetNextContextPosition(LogicalDirection.Forward))
            {
                if (p.GetPointerContext(LogicalDirection.Forward) == TextPointerContext.Text)
                {
                    string runText = p.GetTextInRun(LogicalDirection.Forward);
                    int idx = runText.IndexOf(query, Comparison);
                    if (idx >= 0)
                    {
                        var matchStart = p.GetPositionAtOffset(idx);
                        var matchEnd = matchStart?.GetPositionAtOffset(query.Length);
                        if (matchStart != null && matchEnd != null)
                            return new TextRange(matchStart, matchEnd);
                    }
                }
            }
            return null;
        }

        private bool FindAndSelect()
        {
            var query = FindBox.Text;
            if (string.IsNullOrEmpty(query))
                return false;

            var start = _editor.Selection.End;
            var match = FindFrom(start, query);
            if (match == null)
            {
                MessageBox.Show(this, Loc.T("NotFound"), Loc.T("FindTitle"), MessageBoxButton.OK, MessageBoxImage.Information);
                return false;
            }

            _editor.Selection.Select(match.Start, match.End);
            _editor.Focus();
            var rect = match.Start.GetCharacterRect(LogicalDirection.Forward);
            _editor.ScrollToVerticalOffset(_editor.VerticalOffset + rect.Top - 100);
            return true;
        }

        private void FindNext_Click(object sender, RoutedEventArgs e) => FindAndSelect();

        private void Replace_Click(object sender, RoutedEventArgs e)
        {
            var query = FindBox.Text;
            if (string.IsNullOrEmpty(query))
                return;

            // If the current selection is already the match, replace it; then find next
            if (!_editor.Selection.IsEmpty
                && string.Equals(_editor.Selection.Text, query, Comparison))
            {
                _editor.Selection.Text = ReplaceBox.Text;
            }
            FindAndSelect();
        }

        private void ReplaceAll_Click(object sender, RoutedEventArgs e)
        {
            var query = FindBox.Text;
            if (string.IsNullOrEmpty(query))
                return;

            int count = 0;
            var p = _editor.Document.ContentStart;
            while (true)
            {
                var match = SearchForward(p, query);
                if (match == null)
                    break;
                match.Text = ReplaceBox.Text;
                p = match.End;
                count++;
                if (count > 100000)
                    break;
            }

            MessageBox.Show(this, string.Format(Loc.T("ReplacedCount"), count),
                Loc.T("FindTitle"), MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void Close_Click(object sender, RoutedEventArgs e) => Close();
    }
}
