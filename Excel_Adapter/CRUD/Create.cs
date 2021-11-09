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

namespace BH.Adapter.Excel
{
    public partial class ExcelAdapter
    {
        /***************************************************/
        /**** Public Methods                            ****/
        /***************************************************/

        protected override bool ICreate<T>(IEnumerable<T> objects, ActionConfig actionConfig = null)
        {
            string fileName = m_FileSettings.GetFullFileName();
            XLWorkbook workbook = new XLWorkbook();
            
            if (m_ExcelSettings.NewFile)
                workbook = new XLWorkbook();
            else
                workbook = new XLWorkbook(fileName);

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

            ApplyStyles(workbook);
            ApplyProperties(workbook);
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
            int sheetnum = workbook.Worksheets.Count();
            if (table.Name == null || table.Name == "")
                table.Data.TableName = "sheet " + workbook.Worksheets.Count();
            else
                table.Data.TableName = table.Name;
            if (WorksheetExists(workbook, table.Name))
                table.Data.TableName += sheetnum;

            workbook.Worksheets.Add(table.Data);
            
            return true;
        }

        /***************************************************/

        private bool WorksheetExists(IXLWorkbook workbook, string name)
        {
            foreach(IXLWorksheet worksheet in workbook.Worksheets)
            {
                if (worksheet.Name == name)
                    return true;
            }
            return false;
        }

        /***************************************************/

        private void ApplyStyles(XLWorkbook workbook)
        {
            ApplyWorkbookStyle(workbook);

            foreach (IXLWorksheet sheet in workbook.Worksheets)
            {
                ApplyWorksheetStyle(sheet);

                foreach (IXLTable table in sheet.Tables)
                    ApplyTableStyle(table);  
            }
        }

        /***************************************************/

        private void ApplyProperties(IXLWorkbook workbook)
        {
            workbook.Properties.Author = m_ExcelSettings.WorkbookProperties.Author;
            workbook.Properties.Title = m_ExcelSettings.WorkbookProperties.Title;
            workbook.Properties.Subject = m_ExcelSettings.WorkbookProperties.Subject;
            workbook.Properties.Category = m_ExcelSettings.WorkbookProperties.Category;
            workbook.Properties.Keywords = m_ExcelSettings.WorkbookProperties.Keywords;
            workbook.Properties.Comments = m_ExcelSettings.WorkbookProperties.Comments;
            workbook.Properties.Status = m_ExcelSettings.WorkbookProperties.Status;
            workbook.Properties.LastModifiedBy = m_ExcelSettings.WorkbookProperties.LastModifiedBy;
            workbook.Properties.Company = m_ExcelSettings.WorkbookProperties.Company;
            workbook.Properties.Manager = m_ExcelSettings.WorkbookProperties.Manager;
        }

        /***************************************************/

        private void ApplyWorkbookStyle(XLWorkbook workbook)
        {

        }

        /***************************************************/

        private void ApplyWorksheetStyle(IXLWorksheet worksheet)
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

