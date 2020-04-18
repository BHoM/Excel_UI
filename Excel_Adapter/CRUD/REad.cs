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

using BH.oM.Adapter;
using BH.oM.Base;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using BH.Engine.Adapter;
using ClosedXML.Excel;
using BH.oM.Data.Collections;
using System.Data;

namespace BH.Adapter.ExcelAdapter
{
    public partial class ExcelAdapter
    {
        protected override IEnumerable<IBHoMObject> IRead(Type type, IList ids, ActionConfig actionConfig = null)
        {
            return Read();
        }

        private IEnumerable<IBHoMObject> Read()
        {
            XLWorkbook workbook = new XLWorkbook(_fileSettings.GetFullFileName());
            return ReadExcelFile(workbook);
        }
        /***************************************************/
        /**** Private Methods                           ****/
        /***************************************************/
        private List<IBHoMObject> ReadExcelFile(XLWorkbook workbook, List<string> ids = null)
        {
            List<IBHoMObject> objects = new List<IBHoMObject>();
            foreach (IXLWorksheet worksheet in Worksheets(workbook))
            {
                IXLRange range = Range(worksheet);
                if(range == null)
                {
                    Engine.Reflection.Compute.RecordError("Range provided is not in the correct format for and xlsx file");
                    return objects;
                }
                List<DataColumn> columns = new List<DataColumn>();
                foreach (IXLRangeColumn column in range.Columns())
                {
                    columns.Add(new DataColumn(column.ColumnLetter(), typeof(object)));
                }
                DataTable table = new DataTable();
                table.Columns.AddRange(columns.ToArray());

                foreach (IXLRangeRow row in range.Rows())
                {
                    List<object> dataRow = new List<object>();
                    foreach (IXLRangeColumn column in range.Columns())
                        dataRow.Add(worksheet.Cell(row.RowNumber(),column.ColumnNumber()).GetValue<object>());
                    table.Rows.Add(dataRow.ToArray());
                }
                objects.Add(new Table { Data = table, Name = worksheet.Name });
            }
            return objects;
        }
        /***************************************************/
        private IXLRange Range(IXLWorksheet worksheet)
        {
            if (_excelSettings.Range != null)
            {

                return worksheet.Range(_excelSettings.Range);
            }
            return worksheet.Range(worksheet.FirstCellUsed().Address,worksheet.LastCellUsed().Address);
        }
        /***************************************************/
        private List<IXLWorksheet> Worksheets(XLWorkbook workbook)
        {
            if(_excelSettings.Worksheets != null)
            {
                List<IXLWorksheet> sheets = new List<IXLWorksheet>();
                foreach(string wsName in _excelSettings.Worksheets)
                {
                    if (workbook.Worksheet(wsName) != null)
                        sheets.Add(workbook.Worksheet(wsName));
                }
                return sheets;
            }
            return workbook.Worksheets.ToList();
        }
    }
}

