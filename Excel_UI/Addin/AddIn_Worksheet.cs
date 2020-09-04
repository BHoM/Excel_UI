/*
 * This file is part of the Buildings and Habitats object Model (BHoM)
 * Copyright (c) 2015 - 2020, the respective contributors. All rights reserved.
 *
 * Each contributor holds copyright over their respective contributions.
 * The project versioning (Git) records all such contribution source information.
 *
 *
 * The BHoM is free software: you can redistribute it and/or modify
 * it under the terms of the GNU Lesser General Public License as published by
 * the Free Software Foundation, either version 3.0 of the License, or
 * (at your option) any later version.
 *
 * The BHoM is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
 * GNU Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public License
 * along with this code. If not, see <https://www.gnu.org/licenses/lgpl-3.0.html>.
 */

using System;
using System.IO;
using System.Reflection;
using System.Linq;
using ExcelDna.Integration;
using System.Collections.Generic;
using System.Collections;
using System.Linq.Expressions;
using System.Drawing;
using System.Xml;
using BH.oM.UI;
using BH.Engine.Base;
using BH.Engine.Serialiser;
using NetOffice.ExcelApi.Enums;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;

namespace BH.UI.Excel
{
    public partial class AddIn : IExcelAddIn
    {
        /*******************************************/
        /**** Methods                           ****/
        /*******************************************/

        public static void SaveData(string name, string content, bool replaceExisting = false)
        {
            Application app = ExcelDnaUtil.Application as Application;
            Workbook workbook = app.ActiveWorkbook;

            if (replaceExisting)
            {
                foreach (CustomXMLPart part in workbook.CustomXMLParts.SelectByNamespace($"BH.UI.Excel.{name}").OfType<CustomXMLPart>())
                    part.Delete();
            }

            string xmlString = $"<{name} xmlns=\"BH.UI.Excel.{name}\">{content}</{name}>";
            CustomXMLPart employeeXMLPart = workbook.CustomXMLParts.Add(xmlString);
        }

        /*******************************************/

        public static List<string> ReadData(string name)
        {
            Application app = ExcelDnaUtil.Application as Application;
            Workbook workbook = app.ActiveWorkbook;


            List<CustomXMLPart> parts = workbook.CustomXMLParts.SelectByNamespace($"BH.UI.Excel.{name}").OfType<CustomXMLPart>().ToList();
            return parts.SelectMany(x => x.SelectNodes("/").OfType<CustomXMLNode>()).Select(x => x.Text.Trim()).ToList();
        }

        /*******************************************/

        public static NetOffice.ExcelApi.Worksheet Sheet(string name, bool addIfMissing = true, bool isHidden = false)
        {
            // Look for the sheet in the active workbook
            NetOffice.ExcelApi.Application app = NetOffice.ExcelApi.Application.GetActiveInstance();
            NetOffice.ExcelApi.Workbook workbook = app.ActiveWorkbook;
            NetOffice.ExcelApi.Worksheet sheet = null;
            if (workbook.Sheets.OfType<NetOffice.ExcelApi.Worksheet>().Any(x => x.Name == name))
                sheet = workbook.Sheets[name] as NetOffice.ExcelApi.Worksheet;

            // If sheet doesn't exist, create it if requested
            if (sheet == null && addIfMissing)
            {
                sheet = workbook.Sheets.Add() as NetOffice.ExcelApi.Worksheet;
                sheet.Name = name;

                if (isHidden)
                    sheet.Visible = NetOffice.ExcelApi.Enums.XlSheetVisibility.xlSheetHidden;
            }

            // Return the sheet
            return sheet;
        }

        /*******************************************/

        public static ExcelReference RunningCell()
        {
            return XlCall.Excel(XlCall.xlfCaller) as ExcelReference;
        }

        /*******************************************/

        public static ExcelReference CurrentSelection()
        {
            return XlCall.Excel(XlCall.xlfSelection) as ExcelReference;
        }

        /*******************************************/

        public static void WriteNote(string message, ExcelReference reference = null)
        {
            if (reference == null)
                reference = RunningCell();

            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try
                {
                    XlCall.Excel(XlCall.xlfNote, message, reference);
                }
                catch (XlCallException exception)
                {
                    Engine.Reflection.Compute.RecordError(exception.Message);
                }
            });
        }

        /*******************************************/

        public static void WriteFormula(string formula, ExcelReference reference = null)
        {
            if (reference == null)
                reference = CurrentSelection();

            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try
                {
                    XlCall.Excel(XlCall.xlcFormula, formula, reference);

                    // Let the user fill in the parameters if there is any
                    if (!formula.EndsWith(")"))
                    {
                        Application app = ExcelDnaUtil.Application as Application;
                        if (app != null)
                        {
                            app.SendKeys("{F2}{(}", true);
                            if (System.Windows.Forms.Control.IsKeyLocked(System.Windows.Forms.Keys.NumLock))
                                app.SendKeys("{NUMLOCK}", true);
                        }
                    }
                }
                catch { }
            });
        }

        /*******************************************/
    }
}

