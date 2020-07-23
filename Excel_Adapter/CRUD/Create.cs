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

using BH.oM.Base;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using BH.Engine.Serialiser;
using BH.oM.Adapter;
using BH.Engine.Adapter;
using ClosedXML.Excel;
using BH.oM.Data.Collections;
using System.Data;

namespace BH.Adapter.ExcelAdapter
{
    public partial class ExcelAdapter
    {
        /***************************************************/
        /**** Public Methods                            ****/
        /***************************************************/

        protected override bool ICreate<T>(IEnumerable<T> objects, ActionConfig actionConfig = null)
        {
            string fileName = _fileSettings.GetFullFileName();
            XLWorkbook workbook = new XLWorkbook();
            
            if (_excelSettings.NewFile)
                workbook = new XLWorkbook();
            else
                workbook = new XLWorkbook(fileName);
            
            foreach (T obj in objects)
            {
                //add table to sheet
                if(obj is Table)
                    AddTable(workbook, obj as Table);
            }
            ApplyStyles(workbook);
            ApplyProperties(workbook);
            workbook.SaveAs(fileName); 
            return true;
        }

        /***************************************************/
        /**** Private Methods                           ****/
        /***************************************************/

        private bool AddTable(IXLWorkbook workbook,Table table)
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
            workbook.Properties.Author = _excelSettings.WorkbookProperties.Author;
            workbook.Properties.Title = _excelSettings.WorkbookProperties.Title;
            workbook.Properties.Subject = _excelSettings.WorkbookProperties.Subject;
            workbook.Properties.Category = _excelSettings.WorkbookProperties.Category;
            workbook.Properties.Keywords = _excelSettings.WorkbookProperties.Keywords;
            workbook.Properties.Comments = _excelSettings.WorkbookProperties.Comments;
            workbook.Properties.Status = _excelSettings.WorkbookProperties.Status;
            workbook.Properties.LastModifiedBy = _excelSettings.WorkbookProperties.LastModifiedBy;
            workbook.Properties.Company = _excelSettings.WorkbookProperties.Company;
            workbook.Properties.Manager = _excelSettings.WorkbookProperties.Manager;
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
    }
}

