using System.Collections.Generic;

namespace ExcelProcessor.Interfaces
{
    public interface IXmlMappingService
    {
        (Dictionary<string, int> mappings, int headerRow, int startRow) ReadMappingConfiguration(string xmlPath);
        void ValidateConfiguration(Dictionary<string, int> mappings, int headerRow, int startRow, int totalColumns);
    }
}
