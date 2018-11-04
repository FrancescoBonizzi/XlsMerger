using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using XlsMerger.Domain;
using XlsMerger.Infrastructure;

namespace XlsMerger.Services
{
    public class WorkbookExcelMerger : IExcelMerger
    {
        public async Task<ExcelFileInfos> GetInfos(string excelFilePath)
        {
            return await Task.Run(() =>
            {
                using (var excelApplication = new ExcelApplicationWrapper())
                {
                    var workbook = excelApplication.ExcelApplication.Workbooks.Open(excelFilePath);

                    var excelFileInfos = new ExcelFileInfos(
                        Path.GetFileName(excelFilePath),
                        excelFilePath,
                        workbook.Worksheets.Count);

                    workbook.Close(false, excelFilePath, null);
                    Marshal.ReleaseComObject(workbook);

                    return Task.FromResult(excelFileInfos);
                }
            });
        }

        public async Task MergeFiles(IEnumerable<string> filePaths, string newFilePath)
        {
            await Task.Run(() =>
            {
                using (var excelApplication = new ExcelApplicationWrapper())
                {
                    Workbook newWorkbook = excelApplication.ExcelApplication.Workbooks.Add();

                    // I order it descending because I want the copied sheets in the correct file order,
                    // 1.xlsx, 2.xlsx etc.
                    foreach (var excelFilePath in filePaths)
                    {
                        var thisFileWorkbook = excelApplication.ExcelApplication.Workbooks.Open(excelFilePath);

                        foreach (Worksheet thisWorkbookSheet in thisFileWorkbook.Worksheets)
                        {
                            string newWorkbookSheetName = $"{thisFileWorkbook.Name} - {thisWorkbookSheet.Name}";
                            thisWorkbookSheet.Name = ApplyWorkbookSheetNameRequirements(newWorkbookSheetName);
                            thisWorkbookSheet.Copy(newWorkbook.Worksheets[newWorkbook.Worksheets.Count]);
                        }

                        // I have to discard changes automatically because I edited the source worksheet name
                        thisFileWorkbook.Close(false, excelFilePath);
                        Marshal.ReleaseComObject(thisFileWorkbook);
                    }

                    // Excel interop counts from 1 (and not from zero), therefore, 
                    // removing the second item will cause the third item to take its place!.
                    // To remove the first sheet always empty I have to wait to have other sheets:
                    // you cannot delete the last sheet. 
                    // In this case it is the last one because append all the sheets while merging
                    newWorkbook.Worksheets[newWorkbook.Worksheets.Count].Delete();

                    newWorkbook.SaveAs(newFilePath);
                    Marshal.ReleaseComObject(newWorkbook);
                }
            });
        }

        private string ApplyWorkbookSheetNameRequirements(string newWorkbookSheetName)
        {
            // Excel wants:
            // - No more than 31 characters
            // - No chars: \ / ? * [ ]
            // - No empty name

            newWorkbookSheetName = newWorkbookSheetName
                .Replace("\\", string.Empty)
                .Replace("/", string.Empty)
                .Replace("?", string.Empty)
                .Replace("*", string.Empty)
                .Replace("[", string.Empty)
                .Replace("]", string.Empty);

            if (string.IsNullOrWhiteSpace(newWorkbookSheetName))
                newWorkbookSheetName = $"SheetNoName{Guid.NewGuid().ToString("N")}";

            newWorkbookSheetName = newWorkbookSheetName.Length <= 30
                ? newWorkbookSheetName
                : newWorkbookSheetName.Substring(0, 30);

            return newWorkbookSheetName;
        }
    }
}
