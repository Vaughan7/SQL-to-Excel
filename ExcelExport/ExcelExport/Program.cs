using System.Data;
using Oracle.ManagedDataAccess.Client;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
//using NPOI.OpenXmlFormats.Dml.
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel.Charts;
using NPOI.SS.Util;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Drawing;
using IndexedColors = NPOI.SS.UserModel.IndexedColors;
using DocumentFormat.OpenXml.Office2010.Excel;
using System.Drawing.Text;
using MathNet.Numerics;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ExcelExport
{
    class Program
    {
        private static readonly log4net.ILog logger = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        
        static void Main(string[] args)
        {      

            //Directory.SetCurrentDirectory(AppDomain.CurrentDomain.BaseDirectory); //change directory to exe file location (for task-scheduler)

            QueryManager obj = new QueryManager();//create query manager instance
            obj.connection = obj.ConnectDB();// connect to databse

            var watch = System.Diagnostics.Stopwatch.StartNew(); //start timer
            string currentDateTime = DateTime.Now.ToString("MM/dd/yyyy HH:mm");

            //get queries and execute them                
            string[] sqlFiles = obj.GetQueryFiles();            
            
            foreach (string sqlFile in sqlFiles)
            {
                try
                {
                    string query = File.ReadAllText(sqlFile);
                    string fileName = System.IO.Path.GetFileNameWithoutExtension(sqlFile);
                    ExportToExcel(obj.ExecuteQuery(query), fileName);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error reading file '{sqlFile}': {ex.Message}");
                }
            }

            obj.connection.Dispose();

            //stop timer
            watch.Stop();
            Console.WriteLine($"Execution Time: {watch.Elapsed} ms");

            using (StreamWriter timeWriter = new StreamWriter(AppDomain.CurrentDomain.BaseDirectory + @"..\..\..\..\logs\Timer.log", true))
            {       
                timeWriter.WriteLine($"Current Time: {DateTime.Now.ToString("MM/dd/yyyy HH:mm")} Execution Time: {watch.Elapsed} ms");
            }       

            // Console.Write("write some to close console: ");
            // Console.ReadLine();
        }

        private static void ExportToExcel(DataTable dataTable,string fileName)
        {
            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet1 = workbook.CreateSheet("Sheet1");
        
            //Create Header Stylers
            XSSFCellStyle headerStyle = (XSSFCellStyle)workbook.CreateCellStyle();
        
            //Header Font Styling
            XSSFFont headerFont = (XSSFFont)workbook.CreateFont();
            headerFont.FontHeightInPoints = (short)11;
            headerFont.FontName = "Calibri";
            headerFont.Color = IndexedColors.White.Index;
            headerFont.IsBold = false;
            headerFont.IsItalic = false;
        
            headerStyle.SetFont(headerFont);
        
            //Header Background Color Styling
            byte[] headerColor = new byte[] { 68, 114, 196 };
            headerStyle.SetFillForegroundColor(new XSSFColor(headerColor));
            headerStyle.FillPattern = FillPattern.SolidForeground;
        
            //Header Border Styling
            headerStyle.BorderBottom = BorderStyle.Thin;
            headerStyle.BorderTop = BorderStyle.Thin;
            headerStyle.BorderLeft = BorderStyle.Thin;
            headerStyle.BorderRight = BorderStyle.Thin;
        
            // Create header row
            IRow headerRow = sheet1.CreateRow(0);
            ICell headerTempCell;
        
            for (int i = 0; i < dataTable.Columns.Count; i++)
            {
                headerRow.CreateCell(i).SetCellValue(dataTable.Columns[i].ColumnName);
                headerTempCell = headerRow.GetCell(i);
                headerTempCell.CellStyle = headerStyle;
            }
            
            //Create Data Stylers
            XSSFCellStyle dataStyle1 = (XSSFCellStyle)workbook.CreateCellStyle();
            XSSFCellStyle dataStyle2 = (XSSFCellStyle)workbook.CreateCellStyle();
        
            //Data Font Styling
            XSSFFont dataFont = (XSSFFont)workbook.CreateFont();
            dataFont.FontHeightInPoints = (short)11;
            dataFont.FontName = "Calibri";
            dataFont.Color = IndexedColors.Black.Index;
            dataFont.IsBold = false;
            dataFont.IsItalic = false;
        
            dataStyle1.SetFont(dataFont);
            dataStyle2.SetFont(dataFont);
        
            //Data Background Color Styling
            byte[] accent1 = new byte[3] { 142, 169, 219 };
            dataStyle1.SetFillForegroundColor(new XSSFColor(accent1));
            dataStyle1.FillPattern = FillPattern.SolidForeground;
        
            byte[] accent2 = new byte[3] { 180, 198, 231 };
            dataStyle2.SetFillForegroundColor(new XSSFColor(accent2));
            dataStyle2.FillPattern = FillPattern.SolidForeground;
        
            ICell dataTempCell;
            for (int i = 0; i < dataTable.Rows.Count; i++)
            {
                IRow dataRow = sheet1.CreateRow(i + 1);
                      
                for (int j = 0; j < dataTable.Columns.Count; j++)
                {
                    dataRow.CreateCell(j).SetCellValue(dataTable.Rows[i][j].ToString());
                    dataTempCell = dataRow.GetCell(j);
                    
                    if (dataTempCell.RowIndex.IsEven())
                        dataTempCell.CellStyle = dataStyle1;                    
                    else
                        dataTempCell.CellStyle = dataStyle2;
                }
            }

            //Resize Columns
            for (int i = 0;i < dataTable.Columns.Count;i++)
            {
                sheet1.AutoSizeColumn(i);
            }        
        
            // Save the workbook to a file
            //using (FileStream fs = new FileStream(@"C:\Users\Klaaste Vaughan\Documents\SQLReport.xlsx", FileMode.Create, FileAccess.Write))
            using (FileStream fs = new FileStream(AppDomain.CurrentDomain.BaseDirectory + @"..\..\..\..\Reports\" + fileName + ".xlsx", FileMode.Create, FileAccess.Write))
            {
                workbook.Write(fs);
            }
            Console.WriteLine("Excel file for "+fileName+ " generated successfully.");
        }
    }
}
