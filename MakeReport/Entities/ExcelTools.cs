using System;
using System.IO;
using MakeReport.Entities.Enums;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;

namespace MakeReport.Entities
{
    class ExcelTools
    {
        /// <summary>
        /// Creates a Report in Excel file
        /// </summary>
        public static void DataTableToExcel(System.Data.DataTable dataTable, string excelPath = null)
        {
            try
            {
                // If the table is empty, throw an exception
                if (dataTable == null || dataTable.Columns.Count == 0)
                    throw new ArgumentException("Table is Null or empty!\n");

                // Load excel
                Application excelApp = new Application();

                // Open and create a workbook
                excelApp.Workbooks.Add();

                // Open a worksheet
                _Worksheet workSheet = excelApp.ActiveSheet;

                // Captures the DataTable columns
                for (var i = 0; i < dataTable.Columns.Count; i++)
                {
                    workSheet.Cells[1, i + 1] = dataTable.Columns[i].ColumnName;
                }

                // Captures the DataTable rows
                for (var i = 0; i < dataTable.Rows.Count; i++)
                {
                    for (var j = 0; j < dataTable.Columns.Count; j++)
                    {
                        workSheet.Cells[i + 2, j + 1] = dataTable.Rows[i][j];
                    }
                }

                // Check the path
                if (!string.IsNullOrEmpty(excelPath))
                {
                    // If the file already exists, Delete to create a new file
                    if (File.Exists(excelPath))
                    {
                        File.Delete(excelPath);
                    }

                    try
                    {
                        workSheet.SaveAs(excelPath);
                        excelApp.Quit();
                    }
                    catch (Exception ex)
                    {
                        throw new ArgumentException("Excel file could not be saved, Check filepath!\n" + ex.Message);
                    }
                }
                else
                { // If no path is specified, open the current file in Excel
                    excelApp.Visible = true;
                }
            }
            catch (Exception ex)
            {
                throw new ArgumentException("Error in Export to Excel \nFile: " + excelPath + "\nError: " + ex.Message);
            }
        }

        /// <summary>
        /// Return a list with values found in the file passed
        /// </summary>
        public static List<string> ExcelColumnToList(Workbook workbook, ExcelColumn excelColumn)
        {
            try
            {
                // Create list that will store values
                List<string> listValues = new List<string>();

                // Open a worksheet
                Worksheet worksheet = workbook.Sheets[1];

                // Instantiates the specified column
                Range column = worksheet.Columns[(int)excelColumn];

                // Get the total rows of the column
                int totalLine = worksheet.UsedRange.Rows.Count;

                // Loop that captures all rows in the column and stores it in the list
                for (int i = 2; i <= totalLine; i++)
                {
                    listValues.Add(column.Cells[i].Value2.ToString());
                }

                return listValues;
            }
            catch (Exception ex)
            {
                throw new Exception("Error capturing values ​​in Excel file \nError: " + ex.Message);
            }
        } 
    }
}
