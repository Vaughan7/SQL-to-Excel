﻿using System.Data;
using Oracle.ManagedDataAccess.Client;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace sqltoexcel
{
    class Program
    {
        static void Main(string[] args)
        {
            // Connection string for database
            string connectionString = "Data Source=192.168.9.5:1521/rocdb.bipa.na;User Id=BIPAIT4;Password=Bipa@321;";
            string query = "SELECT CASE  \r\n\r\n" +
                                "WHEN GROUPING (Divisions) = 1 THEN 'Grand Total'\r\n        " +
                                "ELSE (Divisions)\r\n        END AS Division,        \r\n        \r\n        " +
                                "SUM(CASE WHEN org_cat_ent_cd = '21' THEN Total ELSE 0 END) AS \"21\",\r\n        " +
                                "SUM(CASE WHEN org_cat_ent_cd = 'CY' THEN Total ELSE 0 END) AS \"CY\",\r\n        " +
                                "SUM(CASE WHEN org_cat_ent_cd = 'CC' THEN Total ELSE 0 END) AS \"CC\",\r\n        " +
                                "SUM(CASE WHEN org_cat_ent_cd = 'FOR' THEN Total ELSE 0 END) AS \"FOR\",\r\n        " +
                                "SUM(CASE WHEN org_cat_ent_cd = 'DN' THEN Total ELSE 0 END) AS \"DN\",\r\n        " +
                                "SUM(Total) AS Grand_Total\r\nFROM (\r\n        " +
                                
                            "SELECT  \r\n                " +
                                "major_div_code AS Divisions,\r\n                " +
                                "org_cat_ent_cd,\r\n                " +
                                "COUNT(1) AS Total \r\n        " +
                            "FROM icrs_interface.vw_all_entity \r\n        " +
                            "GROUP BY major_div_code, org_cat_ent_cd) sub\r\n\r\n" +
                            "GROUP BY ROLLUP (Divisions) \r\n" +
                            "ORDER BY Divisions";

            using (OracleConnection connection = new OracleConnection(connectionString))
            {
                try
                {
                    // Test the connection
                    connection.Open();
                    // Check the connection state

                    if (connection.State == ConnectionState.Open)
                    {
                        Console.WriteLine("Connection successful!");
                    }
                    else
                    {
                        Console.WriteLine("Connection failed!");
                        return;
                    }

                    // Execute the query and export the resultset to Excel
                    using (OracleCommand command = new OracleCommand(query, connection))
                    {
                        using (OracleDataAdapter adapter = new OracleDataAdapter(command))
                        {
                            // Create a new DataTable to hold the resultset
                            DataTable dataTable = new DataTable();

                            // Fill the DataTable with the resultset from the adapter
                            adapter.Fill(dataTable);

                            // Export the DataTable to Excel
                            Console.WriteLine("call export");
                            ExportToExcel(dataTable);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error: " + ex.Message);
                }
            }
        }
        private static void ExportToExcel(DataTable dataTable)
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
            using (FileStream fs = new FileStream(@"C:\Users\Klaaste Vaughan\Documents\SQLReport.xlsx", FileMode.Create, FileAccess.Write))
            {
                workbook.Write(fs);
            }
            Console.WriteLine("Excel file generated successfully.");
        }
    }
}