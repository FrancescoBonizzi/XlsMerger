
using Microsoft.Win32;
using System.Collections.Generic;
using System.Windows;
using XlsMerger.Infrastructure;

namespace XlsMerger.Services
{
    public class WPFDialogsManager : IDialogsManager
    {
        public void ShowError(string message)
            => MessageBox.Show(message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);

        public void ShowInformation(string message)
            => MessageBox.Show(message, "Information", MessageBoxButton.OK, MessageBoxImage.Information);

        public string NewFilePath()
        {
            var saveFileDialog = new SaveFileDialog
            {
                Filter = "Excel Files|*.xls;*.xlsx;*.xlsm",
                FileName = "MergeOutput.xlsx"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                return saveFileDialog.FileName;
            }

            return null;
        }

        public IEnumerable<string> SelectExcelFilesToMerge()
        {
            var openFilesDialog = new OpenFileDialog()
            {
                Filter = "Excel Files|*.xls;*.xlsx;*.xlsm",
                Multiselect = true
            };

            if (openFilesDialog.ShowDialog() == true)
            {
                return openFilesDialog.FileNames;
            }

            return null;
        }
    }
}
