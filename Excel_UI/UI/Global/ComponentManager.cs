using BH.oM.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NetOffice.ExcelApi;
using BH.Engine.Serialiser;
using BH.UI.Templates;

namespace BH.UI.Excel.Global
{
    static class ComponentManager
    {
        /*************************************/
        /**** Methods                     ****/
        /*************************************/

        public static void Store(this Caller req, string formula)
        {
            foreach (var existing in GetStored())
            {
                if (formula == existing)
                {
                    return;
                }
            }

            string json = req.Write();
            Application app = null;
            Workbook workbook = null;
            Sheets sheets = null;
            Worksheet sheet = null;
            Range cell = null;
            try
            {
                app = Application.GetActiveInstance();
                workbook = app.ActiveWorkbook;
                sheets = workbook.Sheets;
                try
                {
                    sheet = sheets["BHoM_ComponetRequests"] as Worksheet;
                }
                catch
                {
                    sheet = null;

                }
                if (sheet == null)
                {
                    sheet = sheets.Add() as Worksheet;
                    sheet.Name = "BHoM_ComponetRequests";
                }
                sheet.Visible = NetOffice.ExcelApi.Enums.XlSheetVisibility.xlSheetHidden;
                int row = 0;
                string contents = "";
                do
                {
                    row++;
                    cell = sheet.Cells[row, 3];
                    try
                    {
                        contents = cell.Value as string;
                    }
                    catch { }
                } while (contents != null && contents.Length > 0);

                int c = 0;
                while (c < json.Length)
                {
                    sheet.Cells[row, 1].Value = formula;
                    sheet.Cells[row, 2].Value = req.GetType().Name;
                    cell.Value = json.Substring(c);
                    c += (cell.Value as string).Length;
                    cell = cell.Next;
                }
            }
            finally
            {
                if (app != null)
                    app.Dispose();
                if (sheets != null)
                    sheets.Dispose();
                if (sheet != null)
                    sheet.Dispose();
                if (cell != null)
                    cell.Dispose();
            }
        }

        /*************************************/

        public static void Restore()
        {
            foreach (var req in GetComponents())
            {
                ComponentRestored?.Invoke(null, req);
            }

            // Clear the sheet, it will be repopulated
            Application app = null;
            Workbook workbook = null;
            Sheets sheets = null;
            Worksheet sheet = null;
            Range used = null;
            try
            {
                app = Application.GetActiveInstance();
                workbook = app.ActiveWorkbook;
                sheets = workbook.Sheets;
                try
                {
                    sheet = sheets["BHoM_ComponetRequests"] as Worksheet;

                    used = sheet.UsedRange;
                    used.Clear();
                }
                catch { }
            }
            finally
            {
                if (used != null)
                    used.Dispose();
                if (sheet != null)
                    sheet.Dispose();
                if (sheets != null)
                    sheets.Dispose();
                if (workbook != null)
                    workbook.Dispose();
                if (app != null)
                    app.Dispose();
            }
        }

        /*************************************/
        /**** Private Methods             ****/
        /*************************************/

        private static IEnumerable<string> GetStored()
        {
            List<string> formulas = new List<string>();
            Application app = null;
            Workbook workbook = null;
            Sheets sheets = null;
            Worksheet sheet = null;
            Range cell = null;
            Range next = null;
            Range used = null;
            try
            {
                app = Application.GetActiveInstance();
                workbook = app.ActiveWorkbook;
                sheets = workbook.Sheets;
                try
                {
                    sheet = sheets["BHoM_ComponetRequests"] as Worksheet;

                    used = sheet.UsedRange;
                    foreach (Range row in used.Rows)
                    {
                        string str = "";
                        try
                        {
                            cell = row.Cells[1, 1];
                            str = cell.Value.ToString();
                            next = cell.Next;
                            cell.Dispose();
                            cell = next;
                        }
                        catch { }

                        if (str.Length > 0)
                        {
                            formulas.Add(str);
                        }

                        row.Dispose();
                    }
                }
                catch
                {
                }
            }
            finally
            {
                if (next != null)
                    next.Dispose();
                if (cell != null)
                    cell.Dispose();
                if (used != null)
                    used.Dispose();
                if (sheet != null)
                    sheet.Dispose();
                if (sheets != null)
                    sheets.Dispose();
                if (workbook != null)
                    workbook.Dispose();
                if (app != null)
                    app.Dispose();
            }
            return formulas;
        }

        /*************************************/

        private static Dictionary<string, Tuple<string, string>> GetComponents()
        {
            Dictionary<string, Tuple<string, string>> components = new Dictionary<string, Tuple<string, string>>();

            Application app = null;
            Workbook workbook = null;
            Sheets sheets = null;
            Worksheet sheet = null;
            Range cell = null;
            Range next = null;
            Range used = null;
            try
            {
                app = Application.GetActiveInstance();
                workbook = app.ActiveWorkbook;
                sheets = workbook.Sheets;
                try
                {
                    sheet = sheets["BHoM_ComponetRequests"] as Worksheet;

                    used = sheet.UsedRange;
                    foreach (Range row in used.Rows)
                    {
                        string str = "";
                        string key = "";
                        string callerType = "";
                        try
                        {
                            cell = row.Cells[1, 1];
                            key = cell.Value.ToString();
                            cell = row.Cells[1, 2];
                            callerType = cell.Value.ToString();
                            cell = row.Cells[1, 3];
                            while (cell.Value != null && cell.Value is string && (cell.Value as string).Length > 0)
                            {
                                str += cell.Value;
                                next = cell.Next;
                                cell.Dispose();
                                cell = next;
                            }
                        }
                        catch { }

                        if (str.Length > 0)
                        {
                            components.Add(key,new Tuple<string, string>(callerType, str));
                        }

                        row.Dispose();
                    }
                }
                catch
                {
                }
            }
            finally
            {
                if (next != null)
                    next.Dispose();
                if (cell != null)
                    cell.Dispose();
                if (used != null)
                    used.Dispose();
                if (sheet != null)
                    sheet.Dispose();
                if (sheets != null)
                    sheets.Dispose();
                if (workbook != null)
                    workbook.Dispose();
                if (app != null)
                    app.Dispose();
            }

            return components;
        }

        /*************************************/
        /**** Events                      ****/
        /*************************************/

        public static event EventHandler<KeyValuePair<string, Tuple<string, string>>> ComponentRestored;

        /*************************************/
    }
}
