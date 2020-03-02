using BH.oM.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NetOffice.ExcelApi;
using BH.Engine.Serialiser;

namespace BH.UI.Excel.Global
{
    static class ComponentManager
    {
        /*************************************/
        /**** Methods                     ****/
        /*************************************/

        static public void Store(this ComponentRequest req)
        {

            foreach (var existing in GetComponents())
            {
                if(existing.CallerType == req.CallerType && existing.SelectedItem == req.SelectedItem)
                {
                    return;
                }
            }

            string json = req.ToJson();
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
                } catch
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
                    cell = sheet.Cells[row, 1];
                    try
                    {
                        contents = cell.Value as string;
                    }
                    catch { }
                } while (contents != null && contents.Length > 0);

                int c = 0;
                while (c < json.Length)
                {
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

        static public void Restore()
        {
            foreach (var req in GetComponents())
            {
                ComponentRestored?.Invoke(null, req);
            }
        }

        static private List<ComponentRequest> GetComponents()
        {
            List<string> json = new List<string>();

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
                            json.Add(str);
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
                if (app != null)
                    app.Dispose();
                if (sheets != null)
                    sheets.Dispose();
                if (sheet != null)
                    sheet.Dispose();
                if (cell != null)
                    cell.Dispose();
            }

            List<ComponentRequest> requests = new List<ComponentRequest>();
            foreach (string request in json)
            {
                try
                {
                    ComponentRequest req = BH.Engine.Serialiser.Convert.FromJson(request) as ComponentRequest;
                    if (req != null)
                    {
                        requests.Add(req);
                    }
                }
                catch { }
            }
            return requests;
        }

        /*************************************/
        /**** Events                      ****/
        /*************************************/

        public static event EventHandler<ComponentRequest> ComponentRestored;
    }
}
