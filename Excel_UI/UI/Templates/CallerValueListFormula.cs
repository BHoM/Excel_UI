/*
 * This file is part of the Buildings and Habitats object Model (BHoM)
 * Copyright (c) 2015 - 2018, the respective contributors. All rights reserved.
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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;

namespace BH.UI.Excel.Templates
{
    public abstract class CallerValueListFormula : CallerFormula
    {
        public MultiChoiceCaller MultiChoiceCaller
        {
            get
            {
                return Caller as MultiChoiceCaller;
            }
        }

        public CallerValueListFormula() : base()
        {

        }

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

                app = ExcelDnaUtil.Application as Application;
                workbook = app.ActiveWorkbook;
                sheets = workbook.Sheets;
                names = workbook.Names;

                if (options.Count() > 0)
                {
                    ExcelReference xlref = XlCall.Excel(XlCall.xlfCaller) as ExcelReference;
                    if (xlref != null)
                    {
                        var reftext = XlCall.Excel(XlCall.xlfReftext, xlref, true);
                        cell = app.Range[reftext];
                        worksheet = cell.Worksheet;
                        if (worksheet.Name == "BHoM_ValidationHidden")
                        {
                            ExcelReference target;
                            Caller.DataAccessor.SetDataItem(0, ArrayResizer.Resize(options, out target));
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
                            try
                            {
                                n = names.Item(name);
                            }
                            catch
                            {
                                try
                                {
                                    validation_ws = sheets["BHoM_ValidationHidden"];
                                }
                                catch
                                {
                                    validation_ws = sheets.Add();
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
                                    cell.Value = options.FirstOrDefault();
                                    validation = cell.Validation;
                                    validation.Delete();
                                    validation.Add(XlDVType.xlValidateList, Formula1: $"={name}");
                                    validation.IgnoreBlank = true;
                                }
                                finally
                                {
                                    if (cell != null) Marshal.ReleaseComObject(cell);
                                    if (validation != null) Marshal.ReleaseComObject(validation);
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
                if (app != null) Marshal.ReleaseComObject(app);
                if (validation_ws != null) Marshal.ReleaseComObject(validation_ws);
                if (workbook != null) Marshal.ReleaseComObject(workbook);
                if (worksheet != null) Marshal.ReleaseComObject(worksheet);
                if (sheets != null) Marshal.ReleaseComObject(sheets);
                if (names != null) Marshal.ReleaseComObject(names);
                if (n != null) Marshal.ReleaseComObject(n);
            }
            return success;
        }

        protected abstract List<string> GetChoices();

        public override void FillFormula()
        {
            Application app = null;
            Range cell = null;
            try
            {
                app = ExcelDnaUtil.Application as Application;
                cell = app.Selection as Range;
                cell.Formula = "=" + Function + "()";
            }
            finally
            {
                if (app != null) Marshal.ReleaseComObject(app);
                if (cell != null) Marshal.ReleaseComObject(cell);
            }
        }
    }
}