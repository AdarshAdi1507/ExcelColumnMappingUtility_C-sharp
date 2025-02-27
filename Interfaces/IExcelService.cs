using System.Collections.Generic;

namespace ExcelProcessor.Interfaces
{
    public interface IExcelService
    {
        int GetTotalColumns(string filePath, string sheetName);
        void ProcessAndGenerateTeamcenterExcel(
            string sourcePath,
            Dictionary<string, int> mappings,
            string outputPath,
            int headerRow,
            int startRow,
            string sheetName);
    }
}   