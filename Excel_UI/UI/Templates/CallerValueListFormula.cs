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

using BH.Engine.Reflection;
using BH.UI.Templates;
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

        public override bool Run()
        {
            var options = GetChoices().ToArray();

            Application app = null;
            Range cell = null;
            Worksheet validation_ws = null;
            Worksheet worksheet = null;
            Workbook workbook = null;
            Sheets sheets = null;
            Names names = null;
            Name n = null;

            bool success = false;

            try
            {
                var name = $"RANGE_{Function}__";

                app = Application.GetActiveInstance();
                workbook = app.ActiveWorkbook;
                sheets = workbook.Sheets;
                names = workbook.Names;

                if (options.Count() > 0)
                {
                    ExcelReference xlref = XlCall.Excel(XlCall.xlfCaller) as ExcelReference;
                    if (xlref != null)
                    {
                        var reftext = XlCall.Excel(XlCall.xlfReftext, xlref, true);
                        cell = app.Range(reftext);
                        worksheet = cell.Worksheet;
                        if (worksheet.Name == "BHoM_ValidationHidden")
                        {
                            ExcelReference target = xlref;
                            if((xlref.ColumnLast - xlref.ColumnFirst) + 1 != options.Length)
                            {
                                target = new ExcelReference(xlref.RowFirst, xlref.RowFirst, xlref.ColumnFirst, options.Length - 1, xlref.SheetId);
                            }
                            var opt_array = new object[1, options.Length];
                            for (int i = 0; i < options.Length; i++)
                            {
                                opt_array[0, i] = options[i];
                            }
                            Caller.DataAccessor.SetDataItem(0, ArrayResizer.Resize(opt_array));
                            ExcelAsyncUtil.QueueAsMacro(() =>
                            {
                                try
                                {
                                    XlCall.Excel(XlCall.xlcDefineName, name, target);
                                }
                                catch { }
                            });
                            success = true;
                        }
                        else
                        {
                            n = names.FirstOrDefault(nam => nam.Name == name);
                            if (n == null)
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
                                names.Add(name, listcell);
                            }

                            ExcelAsyncUtil.QueueAsMacro(() =>
                            {
                                Validation validation = null;
                                try
                                {
                                    app = Application.GetActiveInstance();
                                    cell = app.Range(reftext);
                                    cell.Value = options.FirstOrDefault();
                                    validation = cell.Validation;
                                    validation.Delete();
                                    validation.Add(XlDVType.xlValidateList, null, null, $"={name}");
                                    validation.IgnoreBlank = true;
                                }
                                finally
                                {
                                    if (app != null)
                                        app.Dispose();
                                    if (cell != null)
                                        cell.Dispose();
                                    if (validation != null)
                                        validation.Dispose();
                                }
                            });

                            Caller.DataAccessor.SetDataItem(0, "");

                            success = true;
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Compute.RecordError(e.GetType().ToText() + ": " + e.Message);
            }
            finally
            {
                if (app != null)
                    app.Dispose();
                if (validation_ws != null)
                    validation_ws.Dispose();
                if (workbook != null)
                    workbook.Dispose();
                if (worksheet != null)
                    worksheet.Dispose();
                if (sheets != null)
                    sheets.Dispose();
                if (names != null)
                    names.Dispose();
                if (cell != null)
                    cell.Dispose();
                if (n != null)
                    n.Dispose();
            }
            return success;
        }

        /*******************************************/

        public override void FillFormula()
        {
            Application app = null;
            Range cell = null;
            try
            {
                app = Application.GetActiveInstance();
                cell = app.Selection as Range;
                cell.Formula = "=" + Function + "()";
            }
            finally
            {
                if (app != null)
                    app.Dispose();
                if (cell != null)
                    cell.Dispose();
            }
        }

        /*******************************************/
        /**** Private Methods                   ****/
        /*******************************************/

        protected abstract List<string> GetChoices();

        /*******************************************/
    }
}
