/*
 * This file is part of the Buildings and Habitats object Model (BHoM)
 * Copyright (c) 2015 - 2024, the respective contributors. All rights reserved.
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

using BH.Engine.Base;
using BH.oM.Base;
using BH.UI.Base;
using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using BH.Engine.Serialiser;

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
        /**** Override Methods                  ****/
        /*******************************************/

        public override object Run(object[] inputs)
        {
            // Collect the list of options
            List<string> names = MultiChoiceCaller.GetChoiceNames();
            List<string> choices = MultiChoiceCaller.Choices.Select((o, i) =>
            {
                if (o is IObject)
                    return $"{names[i]} [{AddIn.IAddObject(o)}]";
                else
                    return names[i];
            }).ToList();

            // Create the dropdown in the cell
            bool success = false;
            try
            {
                if (choices.Count > 0)
                {
                    ExcelReference xlref = AddIn.RunningCell();
                    if (xlref != null)
                    {
                        ExcelAsyncUtil.QueueAsMacro(() =>
                        {
                            Application app = ExcelDnaUtil.Application as Application;
                            string reftext = XlCall.Excel(XlCall.xlfReftext, xlref, true) as string;
                            Range cell = app.Range[reftext];
                            cell.Value = choices.FirstOrDefault();

                            string formula = GetChoicesFormula(MultiChoiceCaller.SelectedItem.ToJson(), choices);

                            Validation validation = cell.Validation;
                            validation.Delete();
                            validation.Add(XlDVType.xlValidateList, XlDVAlertStyle.xlValidAlertWarning, XlFormatConditionOperator.xlBetween, formula, Type.Missing);
                            validation.InCellDropdown = true;
                            validation.IgnoreBlank = true;

                            // Log usage
                            Workbook workbook = app.ActiveWorkbook;
                            Engine.UI.Compute.LogUsage("Excel", app?.Version, InstanceId, Caller.GetType().Name, Caller.SelectedItem, null, AddIn.WorkbookId(workbook), workbook.FullName);
                        });

                        m_DataAccessor.SetDataItem(0, "");
                        success = true;
                    }
                }
            }
            catch (Exception e)
            {
                Engine.Base.Compute.RecordError(e.GetType().ToText() + ": " + e.Message);
            }
            return success;
        }

        /*******************************************/

        protected override void Fill(ExcelReference cell)
        {
            AddIn.WriteFormula("=" + Function + "()", cell);
        }


        /*******************************************/
        /**** Private Methods                   ****/
        /*******************************************/

        protected string GetChoicesFormula(string collectionName, List<string> choices)
        {
            // Get the data sheet
            Worksheet sheet = AddIn.Sheet("BHoM_ChoicesHidden", true, true);
            if (sheet == null)
                return string.Join(",", choices);

            // Try to find the list of choices in the spreadsheet
            int i = 0;
            while (i++ < 1000) // Just for safety
            {
                try
                {
                    string name = sheet.Cells[i,1].Value as string;
                    if (string.IsNullOrEmpty(name))
                    {
                        // Need to add the choices here
                        sheet.Cells[i, 1].Value = collectionName;
                        for (int j = 0; j < choices.Count; j++)
                            sheet.Cells[i, j + 2].Value = choices[j];
                        break;
                    } 
                    else if (collectionName == name)
                    {
                        break;
                    }
                }
                catch
                {
                    break;
                }
            }

            // Create the range
            Range range = sheet.Range[sheet.Cells[i, 2], sheet.Cells[i, choices.Count + 1]];
            return $"=BHoM_ChoicesHidden!{range.Address}";
        }

        /*******************************************/
    }
}




