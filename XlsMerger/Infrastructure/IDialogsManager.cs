using System.Collections.Generic;
using XlsMerger.Domain;

namespace XlsMerger.Infrastructure
{
    public interface IDialogsManager
    {
        IEnumerable<string> SelectExcelFilesToMerge();
        string NewFilePath();
        void ShowInformation(string message);
        void ShowError(string message);
    }
}
