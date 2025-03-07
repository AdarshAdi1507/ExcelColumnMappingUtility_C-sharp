using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using ExcelProcessor.Interfaces;

namespace ExcelProcessor.Services
{
    public class OutputMapping
    {
        public string Filename { get; set; }
        public string Delimiter { get; set; }
        public Dictionary<string, int> ColumnMappings { get; set; } = new Dictionary<string, int>();
    }

    public class XmlMappingService : IXmlMappingService
    {
        private readonly ILogService _logService;

        public XmlMappingService(ILogService logService)
        {
            _logService = logService;
        }

        public (List<OutputMapping> outputMappings, int headerRow, int startRow) ReadMappingConfiguration(string xmlPath)
        {
            try
            {
                var outputMappings = new List<OutputMapping>();
                var doc = new XmlDocument();
                doc.Load(xmlPath);

                // Get HeaderRow and StartRow
                var headerRowNode = doc.SelectSingleNode("//HeaderRow");
                var startRowNode = doc.SelectSingleNode("//StartRow");

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

                // Get all mapping nodes
                var mappingNodes = doc.SelectNodes("//OUTPUTS/Mapping");

                if (mappingNodes == null || mappingNodes.Count == 0)
                {
                    throw new Exception("No output mappings found in the XML configuration");
                }

                foreach (XmlNode mappingNode in mappingNodes)
                {
                    var outputMapping = new OutputMapping();

                    // Get filename attribute
                    var filenameAttr = mappingNode.Attributes["Filename"];
                    if (filenameAttr == null)
                    {
                        throw new Exception("Filename attribute is missing in a Mapping element");
                    }
                    outputMapping.Filename = filenameAttr.Value;

                    // Get delimiter attribute
                    var delimiterAttr = mappingNode.Attributes["delimeter"];
                    outputMapping.Delimiter = delimiterAttr?.Value ?? ","; // Default to comma if not specified

                    // Get column mappings
                    foreach (XmlNode columnNode in mappingNode.ChildNodes)
                    {
                        if (!int.TryParse(columnNode.InnerText, out int columnIndex) || columnIndex <= 0)
                        {
                            throw new Exception($"Invalid column index for {columnNode.Name}: {columnNode.InnerText}. Must be a positive integer.");
                        }
                        outputMapping.ColumnMappings[columnNode.Name] = columnIndex;
                    }

                    outputMappings.Add(outputMapping);
                }

                return (outputMappings, headerRow, startRow);
            }
            catch (Exception ex)
            {
                _logService.LogError("Error reading XML mapping file", ex);
                throw;
            }
        }

        public void ValidateConfiguration(List<OutputMapping> outputMappings, int totalColumns)
        {
            foreach (var mapping in outputMappings)
            {
                var invalidColumns = mapping.ColumnMappings.Where(m => m.Value > totalColumns).ToList();
                if (invalidColumns.Any())
                {
                    throw new Exception($"Following mappings in '{mapping.Filename}' exceed total columns ({totalColumns}): " +
                        $"{string.Join(", ", invalidColumns.Select(x => $"{x.Key}:{x.Value}"))}");
                }
            }
        }
    }
}