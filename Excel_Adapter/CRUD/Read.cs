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
using BH.oM.Excel;
using BH.Engine.Excel;
using BH.oM.Data.Requests;

namespace BH.Adapter.Excel
{
    public partial class ExcelAdapter
    {
        /***************************************************/
        /**** Method Overrides                          ****/
        /***************************************************/

        protected override IEnumerable<IBHoMObject> Read(IRequest request, ActionConfig actionConfig = null)
        {
            //TODO: check if the file and workbook found
            XLWorkbook workbook = new XLWorkbook(m_FileSettings.GetFullFileName());

            if (request is ValuesRequest)
                return ReadExcel(workbook, true);
            else if (request is CellsRequest)
                return ReadExcel(workbook, false);
            else
            {
                BH.Engine.Reflection.Compute.RecordError($"Requests of type {request?.GetType()} are not supported by the Excel adapter.");
                return new List<IBHoMObject>();
            }
        }


        /***************************************************/
        /**** Private Methods                           ****/
        /***************************************************/

        private List<IBHoMObject> ReadExcel(XLWorkbook workbook, bool valuesOnly)
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
                    {
                        if(valuesOnly)
                            dataRow.Add(worksheet.Cell(row.RowNumber(), column.ColumnNumber()).GetValue<object>());
                        else
                            dataRow.Add(BH.Engine.Excel.Create.CellContents(worksheet.Cell(row.RowNumber(), column.ColumnNumber())));
                    }
                        
                    table.Rows.Add(dataRow.ToArray());
                }

                objects.Add(new Table { Data = table, Name = worksheet.Name });
            }

            return objects;
        }
       
        /***************************************************/

        private IXLRange Range(IXLWorksheet worksheet)
        {
            if (m_ExcelSettings.Range != null)
                return worksheet.Range(m_ExcelSettings.Range);
            else
                return worksheet.Range(worksheet.FirstCellUsed().Address, worksheet.LastCellUsed().Address);
        }

        /***************************************************/

        private List<IXLWorksheet> Worksheets(XLWorkbook workbook)
        {
            IEnumerable<IXLWorksheet> result = workbook.Worksheets;
            if (m_ExcelSettings.Worksheets.Count != 0)
                result = result.Where(x => m_ExcelSettings.Worksheets.Contains(x.Name));

            return result.ToList();
        }

        /***************************************************/
    }
}

