using System.Windows;
using MiniWord.Services;

namespace MiniWord
{
    public partial class PreviewWindow : Window
    {
        public PreviewWindow()
        {
            InitializeComponent();
            Title = Loc.T("PreviewTitle");
        }
    }
}
