using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ExcelExport
{
    public class SheetCopier
    {
        private static readonly log4net.ILog logger = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public static void CopySheetToNewWorkbook(string sourceFilePath, string destinationFilePath, string sheetName)
        {
            IWorkbook sourceWorkbook = null;
            IWorkbook destinationWorkbook = new XSSFWorkbook();

            try
            {
                // 1. Open the source workbook and get the sheet
                using (FileStream fsSource = new FileStream(sourceFilePath, FileMode.Open, FileAccess.Read))
                {
                    sourceWorkbook = new XSSFWorkbook(fsSource);
                    ISheet sourceSheet = sourceWorkbook.GetSheet(sheetName);
                    if (sourceSheet == null)
                    {
                        Console.WriteLine($"Sheet '{sheetName}' not found in the source file.");
                        return;
                    }

                    // 2. Create the new sheet in the destination workbook
                    ISheet destinationSheet = destinationWorkbook.CreateSheet(sheetName);

                    // 3. Copy row data, cell styles, and other properties
                    CopySheetData(sourceSheet, destinationSheet);
                }

                // 4. Save the new workbook
                using (FileStream fsDestination = new FileStream(destinationFilePath, FileMode.Create, FileAccess.Write))
                {
                    destinationWorkbook.Write(fsDestination);
                }

                Console.WriteLine($"Successfully copied sheet '{sheetName}' to '{destinationFilePath}'.");
            }
            catch (Exception ex)
            {
                //Console.WriteLine($"An error occurred: {ex.Message}");
                logger.Error(String.Format(ex.Message));
            }
        }

        public static void ImportSheetIntoExistingWorkbook(string sourceFilePath, string destinationFilePath, string sheetName)
        {
            IWorkbook sourceWorkbook = null;
            IWorkbook destinationWorkbook = null;

            try
            {
                // 1. Open the source workbook and get the sheet
                using (FileStream fsSource = new FileStream(sourceFilePath, FileMode.Open, FileAccess.Read))
                {
                    sourceWorkbook = new XSSFWorkbook(fsSource);
                }
                ISheet sourceSheet = sourceWorkbook.GetSheet(sheetName);
                if (sourceSheet == null)
                {
                    Console.WriteLine($"Sheet '{sheetName}' not found in the source file.");
                    return;
                }

                // 2. Open the existing destination workbook
                using (FileStream fsDestination = new FileStream(destinationFilePath, FileMode.Open, FileAccess.ReadWrite))
                {
                    destinationWorkbook = new XSSFWorkbook(fsDestination);

                    // 3. Create the new sheet in the destination workbook
                    ISheet destinationSheet = destinationWorkbook.CreateSheet(sheetName);

                    // Handle potential duplicate sheet names
                    if (destinationWorkbook.GetSheet(sheetName) != null)
                    {
                        int i = 1;
                        string newSheetName;
                        do
                        {
                            newSheetName = $"{sheetName} ({i++})";
                        } while (destinationWorkbook.GetSheet(newSheetName) != null);
                        destinationSheet = destinationWorkbook.CreateSheet(newSheetName);
                    }
                    else
                    {
                        destinationSheet = destinationWorkbook.CreateSheet(sheetName);
                    }

                    // 4. Copy row data, cell styles, and other properties
                    CopySheetData(sourceSheet, destinationSheet);
                }

                // 5. Save the destination workbook after closing the stream and reopening in write mode
                using (FileStream fsDestination = new FileStream(destinationFilePath, FileMode.Create, FileAccess.Write))
                {
                    destinationWorkbook.Write(fsDestination);
                }

                Console.WriteLine($"Successfully copied sheet '{sheetName}' from source file to destination file.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
            finally
            {
                // Clean up resources if necessary
                sourceWorkbook?.Close();
            }
        }

        private static void CopySheetData(ISheet sourceSheet, ISheet destinationSheet)
        {
            // Copy column widths
            for (int i = sourceSheet.FirstRowNum; i < sourceSheet.LastRowNum; i++)
            {
                destinationSheet.SetColumnWidth(i, sourceSheet.GetColumnWidth(i));
            }

            // Copy rows and cells
            for (int rowNum = sourceSheet.FirstRowNum; rowNum <= sourceSheet.LastRowNum; rowNum++)
            {
                IRow sourceRow = sourceSheet.GetRow(rowNum);
                if (sourceRow == null) continue;

                IRow destinationRow = destinationSheet.CreateRow(rowNum);

                for (int cellNum = sourceRow.FirstCellNum; cellNum < sourceRow.LastCellNum; cellNum++)
                {
                    ICell sourceCell = sourceRow.GetCell(cellNum);
                    if (sourceCell == null) continue;

                    ICell destinationCell = destinationRow.CreateCell(cellNum, sourceCell.CellType);

                    // Copy the cell content
                    switch (sourceCell.CellType)
                    {
                        case CellType.String:
                            destinationCell.SetCellValue(sourceCell.StringCellValue);
                            break;
                        case CellType.Numeric:
                            destinationCell.SetCellValue(sourceCell.NumericCellValue);
                            break;
                        case CellType.Boolean:
                            destinationCell.SetCellValue(sourceCell.BooleanCellValue);
                            break;
                        case CellType.Formula:
                            destinationCell.SetCellFormula(sourceCell.CellFormula);
                            break;
                        case CellType.Error:
                            destinationCell.SetCellErrorValue(sourceCell.ErrorCellValue);
                            break;
                            // Copy style and other properties
                    }

                    // Copy cell style (requires copying styles from source workbook to destination)
                    destinationCell.CellStyle = CopyStyle(sourceCell.CellStyle, sourceSheet.Workbook, destinationSheet.Workbook);
                }
            }
        }

        // Helper method to copy cell styles between workbooks
        private static ICellStyle CopyStyle(ICellStyle sourceStyle, IWorkbook sourceWorkbook, IWorkbook destinationWorkbook)
        {
            ICellStyle newStyle = destinationWorkbook.CreateCellStyle();
            newStyle.CloneStyleFrom(sourceStyle);
            return newStyle;
        }
    }
}

