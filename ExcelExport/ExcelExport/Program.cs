using System.Data;
//using Oracle.ManagedDataAccess.Client;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using ExcelExport;

namespace ExcelExport
{
    class Program
    {
        private static readonly log4net.ILog logger = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        
        static void Main(string[] args)
        {

            
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
                        ExportToExcel(obj.ExecuteQuery(query));
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

            using (StreamWriter timeWriter = new StreamWriter(@"..\..\..\..\logs\Timer.log", true))
            {
                
                timeWriter.WriteLine($"Current Time: {DateTime.Now.ToString("MM/dd/yyyy HH:mm")} Execution Time: {watch.Elapsed} ms");
            }


        }

        public static void ExportToExcel(DataTable dataTable)
        {
            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("Sheet1");

            // Create header row
            IRow headerRow = sheet.CreateRow(0);

            for (int i = 0; i < dataTable.Columns.Count; i++)
            {
                headerRow.CreateCell(i).SetCellValue(dataTable.Columns[i].ColumnName);
            }

            // Create data rows
            for (int i = 0; i < dataTable.Rows.Count; i++)
            {
                IRow dataRow = sheet.CreateRow(i + 1);
                for (int j = 0; j < dataTable.Columns.Count; j++)
                {
                    dataRow.CreateCell(j).SetCellValue(dataTable.Rows[i][j].ToString());
                }
            }

            // Save the workbook to a file
            using (FileStream fs = new FileStream(@"..\..\..\..\reports\mySQLReport.xlsx", FileMode.Create, FileAccess.Write))
            {
                workbook.Write(fs);
            }
            Console.WriteLine("Excel file generated successfully.");
        }
    }
}
