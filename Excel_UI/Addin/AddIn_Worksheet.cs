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
using NetOffice.ExcelApi;
using System.Drawing;
using System.Xml;
using BH.oM.UI;
using BH.Engine.Base;
using BH.Engine.Serialiser;
using NetOffice.ExcelApi.Enums;

namespace BH.UI.Excel
{
    public partial class AddIn : IExcelAddIn
    {
        /*******************************************/
        /**** Methods                           ****/
        /*******************************************/

        public static Worksheet Sheet(string name, bool addIfMissing = true, bool isHidden = false)
        {
            // Look for the sheet in the dictionary first
            if (m_SheetReferences.ContainsKey(name) && !m_SheetReferences[name].IsDisposed)
                return m_SheetReferences[name];

            // Look for the sheet in the active workbook
            Application app = Application.GetActiveInstance();
            Workbook workbook = app.ActiveWorkbook;
            Worksheet sheet = null;
            if (workbook.Sheets.OfType<Worksheet>().Any(x => x.Name == name))
                sheet = workbook.Sheets[name] as Worksheet;

            // If sheet doesn't exist, create it if requested
            if (sheet == null && addIfMissing)
            {
                sheet = workbook.Sheets.Add() as Worksheet;
                sheet.Name = name;

                if (isHidden)
                    sheet.Visible = XlSheetVisibility.xlSheetHidden;
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

        public static bool WriteNote(string message, ExcelReference reference = null)
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
            return true;
        }


        /*******************************************/
        /**** Private Fields                    ****/
        /*******************************************/

        private static Dictionary<string, Worksheet> m_SheetReferences = new Dictionary<string, Worksheet>();


        /*******************************************/
    }
}

