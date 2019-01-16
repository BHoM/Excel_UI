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
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;

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

        public CallerValueListFormula(FormulaDataAccessor accessor) : base(accessor)
        {

        }

        public override bool Run()
        {
            var options = GetChoices();

            var app = ExcelDnaUtil.Application as Application;
            var workbook = app.ActiveWorkbook;

            var proj = Project.ForIDs(options);
            try
            {
                ExcelReference xlref = XlCall.Excel(XlCall.xlfCaller) as ExcelReference;
                if (xlref != null)
                {
                    var reftext = XlCall.Excel(XlCall.xlfReftext, xlref, true);
                    ExcelAsyncUtil.QueueAsMacro(() =>
                    {
                        Range cell = app.Range[reftext];
                        cell.Value = options.FirstOrDefault();
                        cell.Validation.Delete();
                        cell.Validation.Add(XlDVType.xlValidateList, Formula1: options.Aggregate((a, b) => $"{a}, {b}"));
                        cell.Validation.IgnoreBlank = true;
                        if (!proj.Empty) proj.SaveData(workbook);
                    });
                    Caller.DataAccessor.SetDataItem(0, "");
                    return true;
                }
            } catch (Exception e)
            {
                Compute.RecordError(e.GetType().ToText() + ": " + e.Message);
            }
            return false;
        }
        
        protected abstract List<string> GetChoices();

        protected override void Caller_ItemSelected(object sender, object e)
        {
            Range cell = Application.Selection as Range;
            cell.Formula = "=" + Function + "()";
        }
    }
}