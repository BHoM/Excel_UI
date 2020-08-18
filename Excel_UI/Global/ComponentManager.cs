using BH.oM.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NetOffice.ExcelApi;
using BH.Engine.Serialiser;
using BH.UI.Base;
using ExcelDna.Integration;

namespace BH.UI.Excel.Global
{
    class ComponentManager : IDisposable
    {
        /*************************************/
        /**** Events                      ****/
        /*************************************/

        public static event EventHandler<KeyValuePair<string, Tuple<string, string>>> ComponentRestored;


        /*************************************/
        /**** Methods                     ****/
        /*************************************/

        public static ComponentManager GetManager(Workbook workbook)
        {
            if (!m_Managers.ContainsKey(workbook.Name))
            {
                m_Managers.Add(workbook.Name, new ComponentManager(workbook));
            }
            return m_Managers[workbook.Name];
        }

        /*************************************/

        public static ComponentManager GetManager(string name)
        {
            if (!m_Managers.ContainsKey(name))
            {
                var workbook = Application.GetActiveInstance().Workbooks[name];
                m_Managers.Add(workbook.Name, new ComponentManager(workbook));
            }
            return m_Managers[name];
        }

        /*************************************/

        public static bool RemoveManager(Workbook workbook)
        {
            try
            {
                if (m_Managers.ContainsKey(workbook.Name))
                {
                    m_Managers[workbook.Name].Dispose();
                    return true;
                }
            }
            catch { }

            return false;
        }

        /*************************************/

        public void Store(Caller req, string formula)
        {
            lock (m_Mutex)
            {
                if (m_Stored.Contains(formula))
                {
                    return;
                }
                string json = req.Write();
                ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    lock (m_Mutex)
                    {
                        if (m_Stored.Contains(formula))
                        {
                            return;
                        }
                        Workbook workbook = m_Workbook;
                        Worksheet sheet = m_Sheet;
                        Range cell = null;
                        if (sheet == null)
                        {
                            sheet = workbook.Sheets.Add() as Worksheet;
                            sheet.Name = "BHoM_ComponetRequests";
                            m_Sheet = sheet;
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

                        m_Stored.Add(formula);
                    }
                });

            }
        }

        /*************************************/

        public void Restore()
        {
            foreach (var restored in GetComponents())
            {
                string key = restored.Key;
                string json = restored.Value.Item2;
                string callerType = restored.Value.Item1;
                if (AddIn.Callers.ContainsKey(callerType))
                {
                    var formula = AddIn.Callers[callerType];
                    if (formula.Caller.Read(json))
                    {
                        if (formula.Function != key)
                        {
                            if (formula.Caller.SelectedItem != null)
                                new UI.Global.ComponentUpgrader(key, formula); // TODO: Look into this, seems weird
                            else
                                return;
                        }
                        formula.Register();
                    }
                }
            }

            if (m_Sheet == null)
                return;

            // Clear the sheet, it will be repopulated
            Range used = null;
            try
            {
                used = m_Sheet.UsedRange;
                used.Clear();
            }
            catch { }
        }

        /*************************************/

        public void Dispose()
        {
            m_Workbook.AfterSaveEvent -= OnWorkbookSaved;
            m_Managers.Remove(m_Name);
        }


        /*************************************/
        /**** Private Methods             ****/
        /*************************************/

        private Dictionary<string, Tuple<string, string>> GetComponents()
        {
            Dictionary<string, Tuple<string, string>> components = new Dictionary<string, Tuple<string, string>>();

            Worksheet sheet = m_Sheet;
            Range cell = null;
            Range used = null;
            if (sheet != null)
            {
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

                        int col = 3;
                        cell = row.Cells[1, col++];
                        while (cell.Value != null && cell.Value is string && (cell.Value as string).Length > 0)
                        {
                            str += cell.Value;
                            cell = row.Cells[1, col++];
                        }
                    }
                    catch { }

                    if (str.Length > 0)
                    {
                        components.Add(key, new Tuple<string, string>(callerType, str));
                    }

                }
            }

            return components;
        }

        /*************************************/

        private void OnWorkbookSaved(bool Success)
        {
            if (m_Workbook.Name != m_Name)
            {
                m_Managers.Add(m_Workbook.Name, this);
                m_Managers.Remove(m_Name);
                m_Name = m_Workbook.Name;
            }
        }

        /*************************************/
        /**** Private Constructors        ****/
        /*************************************/

        ComponentManager(Workbook workbook)
        {
            m_Name = workbook.Name;
            m_Workbook = workbook;
            m_Sheets = workbook.Sheets;
            try
            {
                m_Sheet = m_Sheets["BHoM_ComponetRequests"] as Worksheet;
            }
            catch
            {
                m_Sheet = null;
            }
            m_Workbook.AfterSaveEvent += OnWorkbookSaved;
        }

        /*************************************/
        /**** Private Fields              ****/
        /*************************************/

        private static Dictionary<string, ComponentManager> m_Managers = new Dictionary<string, ComponentManager>();

        private HashSet<string> m_Stored = new HashSet<string>();
        private Workbook m_Workbook;
        private Sheets m_Sheets;
        private Worksheet m_Sheet;
        private string m_Name;
        private object m_Mutex = new object();

        /*************************************/
    }
}
