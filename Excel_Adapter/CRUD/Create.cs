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

using BH.Engine.Adapter;
using BH.oM.Adapter;
using BH.oM.Data.Collections;
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using BH.Engine.Data;
using System.Reflection;
using System.Threading;
using BH.oM.Base;
using BH.Engine.Base;
using BH.oM.Adapters.Excel;

namespace BH.Adapter.Excel
{
    public partial class ExcelAdapter
    {
        /***************************************************/
        /**** Public Methods                            ****/
        /***************************************************/

        protected override bool ICreate<T>(IEnumerable<T> objects, ActionConfig actionConfig = null)
        {
            //TODO: check if the file and workbook found
            string fileName = m_FileSettings.GetFullFileName();
            XLWorkbook workbook = new XLWorkbook(fileName);

            if (actionConfig == null)
            {
                BH.Engine.Reflection.Compute.RecordNote($"{nameof(ExcelPushConfig)} has not been provided, default one is used.");
                actionConfig = new ExcelPushConfig();
            }

            ExcelPushConfig config = actionConfig as ExcelPushConfig;
            if (config == null)
            {
                BH.Engine.Reflection.Compute.RecordError($"Provided {nameof(ActionConfig)} is not {nameof(ExcelPushConfig)}.");
                return false;
            }

            List<Type> objectTypes = objects.Select(x => x.GetType()).Distinct().ToList();
            if (objectTypes.Count != 1)
            {
                string message = "The Excel adapter only allows to push objects of a single type to a table."
                    + "\nRight now you are providing objects of the following types: "
                    + objectTypes.Select(x => x.ToString()).Aggregate((a, b) => a + ", " + b);
                Engine.Reflection.Compute.RecordError(message);
                return false;
            }

            Type type = objectTypes[0];
            if (type == typeof(Table))
            {
                if (objects.Count() == 1)
                    AddTable(workbook, objects.First() as Table);
                else
                {
                    BH.Engine.Reflection.Compute.RecordError("Excel Adapter can push only one table at a time.");
                    return false;
                }
            }
            else
            {
                BH.Engine.Reflection.Compute.RecordError($"Excel Adapter can push only one objects of type {nameof(Table)}.");
                return false;
            }

            ApplyStyles(workbook, config);
            ApplyProperties(workbook, config);
            workbook.SaveAs(fileName);

            //TODO: why is that?
            Thread.Sleep(1000);
            return true;
        }


        /***************************************************/
        /**** Private Methods                           ****/
        /***************************************************/

        private bool AddTable(IXLWorkbook workbook, Table table)
        {
            if (table?.Data == null)
            {
                BH.Engine.Reflection.Compute.RecordError("The input table is null or does not contain a table. Push aborted.");
                return false;
            }

            int sheetCount = workbook.Worksheets.Count();
            string name = table.Name;
            if (string.IsNullOrWhiteSpace(name))
                table.Data.TableName = "Sheet" + sheetCount;

            while (WorksheetExists(workbook, name))
            {
                if (name == "Sheet" + sheetCount)
                    sheetCount++;
                else
                    BH.Engine.Reflection.Compute.RecordWarning($"Worksheet named {name} already exists in the workbook. Default name (sheet 1 etc.) used instead");

                name = "Sheet" + sheetCount;
            }

            DataTable toAdd = table.Data.DeepClone();
            toAdd.TableName = name;

            try
            {
                workbook.Worksheets.Add(toAdd);
                return true;
            }
            catch
            {
                BH.Engine.Reflection.Compute.RecordError("Creation of a new worksheet failed.");
                return false;
            }
        }

        /***************************************************/

        private bool WorksheetExists(IXLWorkbook workbook, string name)
        {
            return workbook.Worksheets.Any(x => x.Name == name);
        }

        /***************************************************/

        private void ApplyStyles(XLWorkbook workbook, ExcelPushConfig config)
        {
            ApplyWorkbookStyle(workbook, config);

            foreach (IXLWorksheet sheet in workbook.Worksheets)
            {
                ApplyWorksheetStyle(sheet, config);

                foreach (IXLTable table in sheet.Tables)
                    ApplyTableStyle(table);  
            }
        }

        /***************************************************/

        private void ApplyProperties(IXLWorkbook workbook, ExcelPushConfig config)
        {
            workbook.Properties.Author = config.WorkbookProperties.Author;
            workbook.Properties.Title = config.WorkbookProperties.Title;
            workbook.Properties.Subject = config.WorkbookProperties.Subject;
            workbook.Properties.Category = config.WorkbookProperties.Category;
            workbook.Properties.Keywords = config.WorkbookProperties.Keywords;
            workbook.Properties.Comments = config.WorkbookProperties.Comments;
            workbook.Properties.Status = config.WorkbookProperties.Status;
            workbook.Properties.LastModifiedBy = config.WorkbookProperties.LastModifiedBy;
            workbook.Properties.Company = config.WorkbookProperties.Company;
            workbook.Properties.Manager = config.WorkbookProperties.Manager;
        }

        /***************************************************/

        private void ApplyWorkbookStyle(XLWorkbook workbook, ExcelPushConfig config)
        {

        }

        /***************************************************/

        private void ApplyWorksheetStyle(IXLWorksheet worksheet, ExcelPushConfig config)
        {

        }

        /***************************************************/

        private void ApplyTableStyle(IXLTable table)
        {
            table.Theme = XLTableTheme.None;
        }

        /***************************************************/
    }
}

