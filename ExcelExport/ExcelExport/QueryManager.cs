//using Microsoft.Office.Interop.Excel;
using Oracle.ManagedDataAccess.Client;
// using System;
// using System.Collections.Generic;
// using System.Linq;
// using System.Text;
// using System.Threading.Tasks;

using System.Data;
using Newtonsoft.Json;
using System.Linq.Expressions;
// using Microsoft.Extensions.Configuration;
// using System.Xml.Serialization;
// using System.Drawing;
// using sqltoexcel;


namespace ExcelExport
{
    internal class QueryManager
    {

        internal OracleConnection? connection;
        private static readonly log4net.ILog logger = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        internal QueryManager()
        {
 
        }



        internal OracleConnection ConnectDB()
        {
            //string connectionString = "";

            //get config
            try
            {
                string jsonString = File.ReadAllText("../../../../config/config.json");

                dynamic? jsonObject = JsonConvert.DeserializeObject(jsonString);

                string connectionString = jsonObject.ConnectionString;
                 
                connection = new OracleConnection(connectionString);


                // Test the connection
                connection.Open();
                // Check the connection state
                if (connection.State == ConnectionState.Open)
                    {Console.WriteLine("Connection successful!");
                    }

            }
            catch(FileNotFoundException)
            {
                //log error ***
                //terminate program ***
            }catch (NullReferenceException){
                //null error
            }catch (Exception ex){
                    Console.WriteLine($"Failed to connect to database: {ex.Message}");
                    logger.Error(String.Format(ex.Message));
            }

            return connection;
        }


       
        public string[] GetQueryFiles()
        {

//add try catch ***
            string folderPath = "../../../../Queries";

            // Get all SQL files from the specified folder
            return Directory.GetFiles(folderPath, "*.sql");
        }


        internal DataTable GetDataTable()
        {

            string folderPath = "../../../../Queries";

            // Get all SQL files from the specified folder
            string[] sqlFiles = Directory.GetFiles(folderPath, "*.sql");

            foreach (string sqlFile in sqlFiles)
            {
                try
                {
                    string query = File.ReadAllText(sqlFile);
                    //Console.WriteLine($"Query from file '{sqlFile}':\n{query}\n");
                    return ExecuteQuery(query);

                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error reading file '{sqlFile}': {ex.Message}");
                    //log to file  ***
                    ///stop program ***
                }
            }
            return new DataTable();
        }


        /** COMING SOON **/
        internal void GetViews()
        {

        }


        internal DataTable ExecuteQuery(string query)
        {
          
            DataTable dataTable = new DataTable();

            try
            {
                OracleCommand command = new OracleCommand(query, connection);

                    using (OracleDataAdapter adapter = new OracleDataAdapter(command))
                    {
                        // Fill the DataTable with the result-set from the adapter
                        adapter.Fill(dataTable);
                    }
            }


            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);

                using (StreamWriter errorWriter = new StreamWriter(@"..\..\..\..\logs\ErrorLog.log", true))
                {
                    errorWriter.WriteLine("Error: " + ex.ToString());
                    //errorWriter.WriteLine("Current Time: " + currentDateTime);
                    logger.Error(String.Format(ex.Message));
                }
            }

            return dataTable;

        }

        /**
            dispose of all objects
        */
       /*  internal void Dispose()
        {

        } */

    }
}
