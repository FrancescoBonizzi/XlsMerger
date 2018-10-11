using GalaSoft.MvvmLight;
using GalaSoft.MvvmLight.CommandWpf;
using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using XlsMerger.Domain;
using XlsMerger.Infrastructure;

namespace XlsMerger.ViewModels
{
    public class MainWindowViewModel : ViewModelBase
    {
        public ObservableCollection<ExcelFileInfos> FilesToMerge { get; set; }

        private readonly IDialogsManager _dialogManager;
        private readonly IExcelMerger _excelMerger;
        
        public ICommand SelectFilesToMergeCommand { get; private set; }
        public ICommand SelectOutputFileCommand { get; private set; }
        public ICommand MergeCommand { get; private set; }
        public ICommand RestartCommand { get; private set; }
        
        public MainWindowViewModel(
            IDialogsManager dialogsManager,
            IExcelMerger excelMerger)
        {
            FilesToMerge = new ObservableCollection<ExcelFileInfos>();
            _dialogManager = dialogsManager ?? throw new ArgumentNullException(nameof(dialogsManager));
            _excelMerger = excelMerger ?? throw new ArgumentNullException(nameof(excelMerger));

            SelectFilesToMergeCommand = new RelayCommand(
                async () =>
                {
                    var filesToMergePaths = _dialogManager.SelectExcelFilesToMerge();
                    if (filesToMergePaths == null)
                        return;

                    FilesToMerge.Clear();

                    ProgressBarVisibility = Visibility.Visible;
                    foreach (var fileToMergePath in filesToMergePaths.OrderBy(f => f))
                        FilesToMerge.Add(await _excelMerger.GetInfos(fileToMergePath));
                    ProgressBarVisibility = Visibility.Collapsed;
                });

            SelectOutputFileCommand = new RelayCommand(
                () =>
                {
                    var outputFilePath = _dialogManager.NewFilePath();
                    if (outputFilePath == null)
                        return;

                    NewFilePath = outputFilePath;
                });

            MergeCommand = new RelayCommand(
               async () =>
                {
                    try
                    {
                        ProgressBarVisibility = Visibility.Visible;
                        await _excelMerger.MergeFiles(FilesToMerge.Select(f => f.FilePath), NewFilePath);
                        ProgressBarVisibility = Visibility.Hidden;
                        _dialogManager.ShowInformation($"Operation completed");
                    }
                    catch (Exception ex)
                    {
                        _dialogManager.ShowError(ex.Message);
                    }
                },
                () => FilesToMerge.Any() && !string.IsNullOrWhiteSpace(NewFilePath));

            RestartCommand = new RelayCommand(
                () =>
                {
                    FilesToMerge.Clear();
                    NewFilePath = string.Empty;
                    ProgressBarVisibility = Visibility.Collapsed;
                },
                () => FilesToMerge.Any() || !string.IsNullOrWhiteSpace(NewFilePath));
        }

        private string _newFilePath;
        public string NewFilePath
        {
            get => _newFilePath;
            set
            {
                if (_newFilePath == value)
                    return;

                _newFilePath = value;
                RaisePropertyChanged(nameof(NewFilePath));
            }
        }

        private Visibility _progressBarVisibility = Visibility.Collapsed;
        public Visibility ProgressBarVisibility
        {
            get => _progressBarVisibility;
            set
            {
                if (_progressBarVisibility == value)
                    return;

                _progressBarVisibility = value;
                RaisePropertyChanged(nameof(ProgressBarVisibility));
            }
        }

    }
}
