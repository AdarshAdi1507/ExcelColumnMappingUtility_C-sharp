using ExcelProcessor.Services;
namespace ExcelProcessor.Interfaces { 
public interface IExcelService
{
    int GetTotalColumns(string filePath, string sheetName);
    void ProcessAndGenerateOutputFiles(
        string sourcePath,
        List<OutputMapping> outputMappings,
        int headerRow,
        int startRow,
        string sheetName);
}
}