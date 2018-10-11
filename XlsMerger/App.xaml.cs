using System.Windows;
using XlsMerger.Infrastructure;
using XlsMerger.Services;
using XlsMerger.ViewModels;

namespace XlsMerger
{
    public partial class App : Application
    {
        private IExcelMerger _excelMerger;

        void App_Startup(object sender, StartupEventArgs e)
        {
            var dialogsManager = new WPFDialogsManager();
            _excelMerger = new WorkbookExcelMerger();
            var mainWindowViewModel = new MainWindowViewModel(
                dialogsManager,
                _excelMerger);
            var window = new MainWindow(mainWindowViewModel);
            window.Show();
        }
    }
}
