using System.Windows;
using XlsMerger.ViewModels;

namespace XlsMerger
{
    public partial class MainWindow : Window
    {
        public MainWindow(MainWindowViewModel viewModel)
        {
            InitializeComponent();
            DataContext = viewModel;
        }
    }
}
