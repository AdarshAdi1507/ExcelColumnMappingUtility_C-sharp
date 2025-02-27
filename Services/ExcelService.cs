using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using ExcelProcessor.Interfaces;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Security.Policy;

namespace ExcelProcessor.Services
{
    public class ExcelService : IExcelService
    {
        private readonly ILogService _logService;

        public ExcelService(ILogService logService)
        {
            _logService = logService;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        public int GetTotalColumns(string filePath, string sheetName)
        {
            try
            {
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    var worksheet = package.Workbook.Worksheets[sheetName];
                    return worksheet.Dimension.Columns;
                }
            }
            catch (Exception ex)
            {
                _logService.LogError("Error getting total columns", ex);
                throw;
            }
        }
        public void ProcessAndGenerateTeamcenterExcel(string sourcePath, Dictionary<string, int> mappings,
    string outputPath,
    int headerRow,
    int startRow,
    string sheetName)
        {
            try
            {
                // Cast the logger to use the new methods
                var logger = _logService as ExcelProcessLogger;
                logger?.StartProcess(sourcePath, sheetName);

                using (var sourcePackage = new ExcelPackage(new FileInfo(sourcePath)))
                using (var targetPackage = new ExcelPackage())
                {
                    var sourceWorksheet = sourcePackage.Workbook.Worksheets[sheetName];
                    var targetWorksheet = targetPackage.Workbook.Worksheets.Add("Sheet1");

                    // Validate input file
                    if (sourceWorksheet.Dimension.Rows < startRow)
                    {
                        throw new Exception($"Source file has fewer rows ({sourceWorksheet.Dimension.Rows}) than the specified StartRow ({startRow})");
                    }

                    // Write headers
                    int col = 1;
                    foreach (var mapping in mappings)
                    {
                        targetWorksheet.Cells[1, col].Value = mapping.Key;
                        col++;
                    }

                    
                    // Copy data according to mapping
                    int targetRow = 2; // Start from row 2 as row 1 has headers
                    for (int sourceRow = startRow; sourceRow <= sourceWorksheet.Dimension.Rows; sourceRow++)
                    {
                        col = 1;  
                        foreach (var mapping in mappings)
                        {
                            // Get value from source based on the mapping
                            var sourceValue = sourceWorksheet.Cells[sourceRow, mapping.Value].Text;
                            // Set value in target worksheet
                            targetWorksheet.Cells[targetRow, col].Value = sourceValue;
                            col++;
                        }
                        targetRow++;
                        logger?.LogRowProcessed();
                    }

                    // After saving the file:
                    targetPackage.SaveAs(new FileInfo(outputPath));
                    logger?.EndProcess(outputPath);
                }
            }
            catch (Exception ex)
            {
                _logService.LogError("Error generating Teamcenter Excel file", ex);
                throw;
            }
        }


        //public void ProcessAndGenerateTeamcenterExcel_BOM(string sourcePath,
        //    Dictionary<string, int> mappings,
        //    string outputPath,
        //    int headerRow,
        //    int startRow,
        //    string sheetName)
        //{
        //    try
        //    {
        //        using (var sourcePackage = new ExcelPackage(new FileInfo(sourcePath)))
        //        using (var targetPackage = new ExcelPackage())
        //        {
        //            var sourceWorksheet = sourcePackage.Workbook.Worksheets[sheetName];
        //            var targetWorksheet = targetPackage.Workbook.Worksheets.Add("Sheet1");

        //            // Validate input file has enough rows
        //            if (sourceWorksheet.Dimension.Rows < startRow)
        //            {
        //                throw new Exception($"Source file has fewer rows ({sourceWorksheet.Dimension.Rows}) than the specified StartRow ({startRow})");
        //            
        //            // Write headers
        //            int col = 1;
        //            foreach (var mapping in mappings)
        //            {
        //                targetWorksheet.Cells[1, col].Value = mapping.Key;
        //                col++;
        //            
        //            // Copy data according to mapping
        //            int targetRow = 2; // Start from row 2 as row 1 has headers
        //            for (int sourceRow = startRow; sourceRow <= sourceWorksheet.Dimension.Rows; sourceRow++)
        //            {
        //                string level_Str = sourceWorksheet.Cells[sourceRow,1].Text
        //                foreach (var mapping in mappings)
        //                
        //                    //if (level_Str == "")
        //                    //{
        //                    //    break ;
        //                    //    continue;
        //                    
        //                    int level = int.Parse(level_Str);
        //                    string currentName = worksheet.Cells[row, 2].Text
        //                    if (currentName == "")
        //                    {
        //                        continue;
        //                    
        //                    string currentRevision = worksheet.Cells[row, 6].Text
        //                    string[] split_space = currentRevision.Split(' 
        //                    string currentRev_Spllitted = split_space[0].Trim
        //                    if (currentRev_Spllitted == "")
        //                    {
        //                        continue;
        //                    
        //                    string currentNameRev = currentName + "~" + currentRev_Spllitted
        //                    // If the current level is greater than 0, generate the combination with its immediate parent
        //                    if (level > 0 && parentDict.ContainsKey(level - 1))
        //                    
        //                        //--------------------------NewlyAdded for BOM
        //                        //--------------------------NewlyAdded for BOM-------------------------------

        //                        // Get the immediate parent (from the previous level)
        //                        string parentName = parentDict[level - 1];

        //                        // Create the combination: Parent~Child (Immediate Parent-Child)
        //                        result.Add(parentName + "~" + currentNameRev);
        //                    }

        //                    // Store the current name as the parent for the next level
        //                    parentDict[level] = currentNameRev;
        //                    has context 
        //                          col = 1
        //                    targetWorksheet.Cells[targetRow, col].Value = 

        //                        sourceWorksheet.Cells[sourceRow, mapping.Value].Text;
        //                    col++;

        //                }
        //                targetRow++;
        //            }

        //            // Save the new Excel file
        //            targetPackage.SaveAs(new FileInfo(outputPath));
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        _logService.LogError("Error generating Teamcenter Excel file", ex);
        //        throw;
        //    }
        //}
    }
}
