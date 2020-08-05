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

using BH.Engine.Excel;
using BH.Engine.Reflection;
using BH.UI.Base;
using ExcelDna.Integration;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using NetOffice.ExcelApi;
using NetOffice.ExcelApi.Enums;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;

namespace BH.UI.Excel.Templates
{
    public abstract class CallerValueListFormula : CallerFormula
    {
        /*******************************************/
        /**** Properties                        ****/
        /*******************************************/

        public MultiChoiceCaller MultiChoiceCaller
        {
            get
            {
                return Caller as MultiChoiceCaller;
            }
        }

        /*******************************************/
        /**** Constructors                      ****/
        /*******************************************/

        public CallerValueListFormula() : base()
        {

        }

        /*******************************************/
        /**** Methods                           ****/
        /*******************************************/

        /*******************************************/
        /**** Private Methods                   ****/
        /*******************************************/
        protected override bool Excecute()
        {
            var options = GetChoices().ToArray();

            Application app = null;
            Range cell = null;
            Worksheet validation_ws = null;
            Worksheet worksheet = null;
            Workbook workbook = null;
            Sheets sheets = null;

            bool success = false;

            try
            {
                var name = $"RANGE_{Function}__";

                app = Application.GetActiveInstance();
                workbook = app.ActiveWorkbook;
                sheets = workbook.Sheets;

                if (options.Count() > 0)
                {
                    ExcelReference xlref = XlCall.Excel(XlCall.xlfCaller) as ExcelReference;
                    if (xlref != null)
                    {
                        string reftext = XlCall.Excel(XlCall.xlfReftext, xlref, true) as string;
                        cell = app.Range(reftext);
                        worksheet = cell.Worksheet;
                        if (worksheet.Name == "BHoM_ValidationHidden")
                        {
                            m_DataAccessor.SetDataItem(0,
                                ArrayResizer.Resize(options, (target) =>
                                {
                                    try
                                    {
                                        XlCall.Excel(XlCall.xlcDefineName, name, target);
                                    }
                                    catch { }
                                })
                            );
                            success = true;
                        }
                        else
                        {
                            ExcelAsyncUtil.QueueAsMacro(() =>
                            {
                                    string prefix = reftext.Substring(0, reftext.LastIndexOf('!') + 1);
                                    var nameDef = XlCall.Excel(XlCall.xlfGetName, prefix + name);
                                    if (nameDef.Equals(ExcelError.ExcelErrorName))
                                    {
                                        try
                                        {
                                            validation_ws = sheets["BHoM_ValidationHidden"] as Worksheet;
                                        }
                                        catch
                                        {
                                            validation_ws = sheets.Add() as Worksheet;
                                            validation_ws.Name = "BHoM_ValidationHidden";
                                        }
                                        validation_ws.Visible = XlSheetVisibility.xlSheetHidden;

                                        int row = 1;
                                        Range listcell = validation_ws.Cells[row, 1];
                                        while (listcell.Value != null)
                                        {
                                            row++;
                                            listcell = validation_ws.Cells[row, 1];
                                        }
                                        listcell.Formula = $"={Function}()";
                                    }

                                    ExcelAsyncUtil.QueueAsMacro(() =>
                                    {
                                        Validation validation = null;
                                            app = Application.GetActiveInstance();
                                            cell = app.Range(reftext);
                                            cell.Value = options.FirstOrDefault();
                                            validation = cell.Validation;
                                            validation.Delete();
                                            validation.Add(XlDVType.xlValidateList, null, null, $"={name}");
                                            validation.IgnoreBlank = true;
                                    });
                            });

                            m_DataAccessor.SetDataItem(0, "");

                            success = true;
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Engine.Reflection.Compute.RecordError(e.GetType().ToText() + ": " + e.Message);
            }
            return success;
        }

        protected override void Fill(oM.Excel.Reference cell)
        {
            var cellcontents = "=" + Function + "()";
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                XlCall.Excel(XlCall.xlcFormula, cellcontents, cell.ToExcel());
            });
        }

        /*******************************************/

        protected abstract List<string> GetChoices();

        /*******************************************/
    }
}
