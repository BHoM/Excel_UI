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
using BH.oM.Base;
using BH.UI.Base;
using ExcelDna.Integration;
using NetOffice.ExcelApi;
using NetOffice.ExcelApi.Enums;
using System;
using System.Collections.Generic;
using System.Linq;

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

        public override object Run(object[] inputs)
        {
            // Collect the list of options
            List<string> names = MultiChoiceCaller.GetChoiceNames();
            string[] options = MultiChoiceCaller.Choices.Select((o, i) =>
            {
                if (o is IObject)
                    return $"{names[i]} [{AddIn.IAddObject(o)}]";
                else
                    return names[i];
            }).ToArray();

            // Create the dropdown in the cell
            bool success = false;
            try
            {
                if (options.Count() > 0)
                {
                    ExcelReference xlref = AddIn.RunningCell();
                    if (xlref != null)
                    {
                        ExcelAsyncUtil.QueueAsMacro(() =>
                        {
                            Application app = Application.GetActiveInstance();
                            string reftext = XlCall.Excel(XlCall.xlfReftext, xlref, true) as string;
                            Range cell = app.Range(reftext);
                            cell.Value = options.FirstOrDefault();

                            Validation validation = cell.Validation;
                            validation.Delete();
                            validation.Add(XlDVType.xlValidateList, null, null, string.Join(",", options));
                            validation.IgnoreBlank = true;
                        });

                        m_DataAccessor.SetDataItem(0, "");
                        success = true;
                    }
                }
            }
            catch (Exception e)
            {
                Engine.Reflection.Compute.RecordError(e.GetType().ToText() + ": " + e.Message);
            }
            return success;
        }


        /*******************************************/
        /**** Private Methods                   ****/
        /*******************************************/

        protected override void Fill(ExcelReference cell)
        {
            var cellcontents = "=" + Function + "()";
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                XlCall.Excel(XlCall.xlcFormula, cellcontents, cell);
            });
        }

        /*******************************************/
    }
}
