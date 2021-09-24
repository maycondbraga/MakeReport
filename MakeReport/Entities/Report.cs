using System;
using System.IO;
using System.Data;
using System.Linq;
using System.Data.SqlClient;
using MakeReport.Entities.Enums;
using System.Collections.Generic;
using System.Configuration;

namespace MakeReport.Entities
{
    class Report
    {
        // Directory to save Excel files
        private static string FilePath = @"C:\Users\Public\Downloads";

        // Query you want to do
        private static string BasicQuery = "SELECT [COLUMNS NAMES] " +
                                           "FROM [TABLE NAME] " +
                                           "WHERE [CONDITIONS]";

        public static string IntermediateQuery = "SELECT [COLUMNS NAMES] " +
                                                 "FROM [TABLE NAME] " +
                                                 "WHERE [CONDITIONS] AND [COLUMN_NAME] IN ({0})";

        /// <summary>
        /// Create simple report based on a query without parameters
        /// </summary>
        public static void BasicReport()
        {
            // Excel file name with current date formated as DayMonthYear
            string filename = @"REPORT_NAME_" + DateTime.Now.ToString("ddMMyyyy") + ".xlsx";

            // Path to save the Excel file
            string fullPath = Path.Combine(FilePath, filename);

            // DataTable that stores the SQL query
            System.Data.DataTable dataTable = new System.Data.DataTable();

            // SQL Server Connection String
            string connectionString = ConfigurationManager.ConnectionStrings["NameDatabase"].ConnectionString;

            // SQL Server Connection
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                // SQL Query Command
                SqlCommand cmd = new SqlCommand(BasicQuery, conn);

                // Open connection with SQL
                conn.Open();

                // Create Data Adapter
                using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                {
                    // Executes the query and returns the result populating the DataTable
                    da.Fill(dataTable);
                }

                // Close connection with SQL
                conn.Close();
            }

            // Method that exports the DataTable as Excel at the specified path
            ExcelTools.DataTableToExcel(dataTable, fullPath);
            GC.Collect();
            GC.WaitForPendingFinalizers();

            Console.WriteLine("\nReport Finished with Success");
        }

        /// <summary>
        /// Creates a report based on a parameterized query
        /// </summary>
        public static void IntermediateReport()
        {
            // If you want a copy of the folders that were loaded by automation, keep true
            bool copy = true;

            try
            {
                // DataTable that stores the SQL query
                System.Data.DataTable dataTable = new System.Data.DataTable();

                List<string> listValues;

                // Gets all file paths in the directory
                string[] path = Directory.GetFiles(FilePath, "*", SearchOption.TopDirectoryOnly);

                if (path.Length != 1)
                {
                    throw new Exception("Leave only the Excel file in the folder\n");
                }

                try
                {
                    // Return a list with values found in the file passed
                    listValues = ExcelTools.ExcelColumnToList(path[0], ExcelColumn.B, true);
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
                catch (Exception ex)
                {
                    throw new Exception(ex.Message);
                }

                try
                {

                    // SQL Server Connection String
                    string connectionString = ConfigurationManager.ConnectionStrings["NameDatabase"].ConnectionString;

                    // SQL Server Connection
                    using (SqlConnection conn = new SqlConnection(connectionString))
                    {

                        // For each item in listValues, an @tag is appended to the paramNames array
                        string[] paramNames = listValues.Select((s, i) => "@tag" + i.ToString()).ToArray();

                        // Format paramNames to SQL
                        string inClause = string.Join(", ", paramNames);

                        // SQL Query Command
                        SqlCommand cmd = new SqlCommand(string.Format(IntermediateQuery, inClause), conn);

                        // For each parameter in paramNames, a value is given in the query
                        for (int i = 0; i < paramNames.Length; i++)
                        {
                            cmd.Parameters.AddWithValue(paramNames[i], listValues[i]);
                        }

                        // Open connection with SQL
                        conn.Open();

                        // Create Data Adapter
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            // Executes the query and returns the result populating the DataTable
                            da.Fill(dataTable);
                        }

                        // Close connection with SQL
                        conn.Close();
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception("Error getting folders from database \nError:" + ex.Message);
                }

                try
                {
                    if (copy)
                    {
                        // Excel file name with current date formated as DayMonthYear
                        string filename = @"REPORT_NAME_" + DateTime.Now.ToString("ddMMyyyy") + ".xlsx";
                        
                        string[] Excelpath = new string[] { FilePath, "Reports" ,  filename };
                        
                        // Path to save the Excel file
                        string fullPath = Path.Combine(Excelpath);

                        // Method that exports the DataTable as Excel at the specified path
                        ExcelTools.DataTableToExcel(dataTable, fullPath);
                        GC.Collect();
                        GC.WaitForPendingFinalizers();

                        Console.WriteLine("\nCopy with folders and id created");
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception(ex.Message);
                }

                Console.WriteLine("\nReport finished with success");
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
    }
}
