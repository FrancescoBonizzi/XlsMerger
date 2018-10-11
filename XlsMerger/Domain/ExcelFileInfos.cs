using System;

namespace XlsMerger.Domain
{
    public class ExcelFileInfos
    {
        public string FileName { get; set; }
        public string FilePath { get; set; }
        public int TabsNumber { get; set; }

        public ExcelFileInfos(string fileName, string filePath, int tabsNumber)
        {
            FileName = fileName ?? throw new ArgumentNullException(nameof(fileName));
            FilePath = filePath ?? throw new ArgumentNullException(nameof(filePath));
            TabsNumber = tabsNumber;
        }
    }
}
