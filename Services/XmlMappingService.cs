using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using ExcelProcessor.Interfaces;

namespace ExcelProcessor.Services
{
    public class XmlMappingService : IXmlMappingService
    {
        private readonly ILogService _logService;

        public XmlMappingService(ILogService logService)
        {
            _logService = logService;
        }

        public (Dictionary<string, int> mappings, int headerRow, int startRow) ReadMappingConfiguration(string xmlPath)
        {
            try
            {
                var mappings = new Dictionary<string, int>();
                var doc = new XmlDocument();
                doc.Load(xmlPath);

                // Get HeaderRow and StartRow
                var headerRowNode = doc.SelectSingleNode("CONFIG/CSV2TCXML/HeaderRow");
                var startRowNode = doc.SelectSingleNode("CONFIG/CSV2TCXML/StartRow");

                if (headerRowNode == null || startRowNode == null)
                {
                    throw new Exception("HeaderRow and StartRow must be specified in the XML configuration");
                }

                if (!int.TryParse(headerRowNode.InnerText, out int headerRow) || headerRow <= 0)
                {
                    throw new Exception("HeaderRow must be a positive integer");
                }

                if (!int.TryParse(startRowNode.InnerText, out int startRow) || startRow <= 0)
                {
                    throw new Exception("StartRow must be a positive integer");
                }

                if (headerRow >= startRow)
                {
                    throw new Exception("HeaderRow must be less than StartRow");
                }

                var mappingNode = doc.SelectSingleNode("CONFIG/CSV2TCXML/ITEMS_INPUT/Mapping");
                if (mappingNode == null)
                {
                    throw new Exception("Invalid mapping XML structure");
                }

                foreach (XmlNode node in mappingNode.ChildNodes)
                {
                    if (!int.TryParse(node.InnerText, out int columnIndex) || columnIndex <= 0)
                    {
                        throw new Exception($"Invalid column index for {node.Name}: {node.InnerText}. Must be a positive integer.");
                    }

                    mappings[node.Name] = columnIndex;
                }

                return (mappings, headerRow, startRow);
            }
            catch (Exception ex)
            {
                _logService.LogError("Error reading XML mapping file", ex);
                throw;
            }
        }

        public void ValidateConfiguration(Dictionary<string, int> mappings, int headerRow, int startRow, int totalColumns)
        {
            var invalidColumns = mappings.Where(m => m.Value > totalColumns).ToList();
            if (invalidColumns.Any())
            {
                throw new Exception($"Following mappings exceed total columns ({totalColumns}): {string.Join(", ", invalidColumns.Select(x => $"{x.Key}:{x.Value}"))}");
            }
        }
    }
}