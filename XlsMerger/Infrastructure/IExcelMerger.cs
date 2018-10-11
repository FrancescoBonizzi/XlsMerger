using System.Collections.Generic;
using System.Threading.Tasks;
using XlsMerger.Domain;

namespace XlsMerger.Infrastructure
{
    public interface IExcelMerger
    {
        Task MergeFiles(IEnumerable<string> filePaths, string newFilePath);
        Task<ExcelFileInfos> GetInfos(string excelFilePath);
    }
}
