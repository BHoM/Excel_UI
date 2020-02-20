using System;
using System.Data;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using BH.oM.Data.Collections;

namespace Excel_Adapter
{
    public static partial class Read
    {
        public static Table ParseExcelSpreadSheet(string path)
        {
            Application xlApp = new Application();
            Workbook xlWorkbook;
            try
            {
                xlWorkbook = xlApp.Workbooks.Open(@path);
            }
            catch
            {
                return null;
            }
            _Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Range x = xlWorksheet.UsedRange;

            int rowCount = xlWorksheet.Cells.Find("*", Missing.Value, Missing.Value, Missing.Value, XlSearchOrder.xlByRows, XlSearchDirection.xlPrevious, false, Missing.Value, Missing.Value).Row;
            int colCount = xlWorksheet.Cells.Find("*", Missing.Value, Missing.Value, Missing.Value, XlSearchOrder.xlByColumns, XlSearchDirection.xlPrevious, false, Missing.Value, Missing.Value).Column;

            Table table = new Table();
            table.Data = new System.Data.DataTable();

            //Creating columns
            for (int i = 1; i < colCount + 1; i++)
                table.Data.Columns.Add(x.Cells[1, i].Value2.ToString(), typeof(string));

            //Filling columns with rows of data
            for (int i = 2; i < rowCount + 1; i++)
            {
                DataRow row = table.Data.NewRow();
                for (int j = 1; j < colCount + 1; j++)
                {
                    string head = table.Data.Columns[j - 1].ColumnName;
                    if (x.Cells[i, j].Value2 != null)
                        row[head] = x.Cells[i, j].Value2.ToString();
                }
                table.Data.Rows.Add(row);
            }

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(x);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);

            return table;
        }
    }
}

