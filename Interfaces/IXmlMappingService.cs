using ExcelProcessor.Services;
namespace ExcelProcessor.Interfaces
{
    public interface IXmlMappingService
    {
        (List<OutputMapping> outputMappings, int headerRow, int startRow) ReadMappingConfiguration(string xmlPath);
        void ValidateConfiguration(List<OutputMapping> outputMappings, int totalColumns);
    }
}