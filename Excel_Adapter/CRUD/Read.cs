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
using BH.oM.Adapters.Excel;
using BH.oM.Base;
using BH.oM.Data.Collections;
using BH.oM.Data.Requests;
using ClosedXML.Excel;
using System.Collections.Generic;
using System.Data;
using System.Linq;

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
                return ReadExcel(workbook, ((ValuesRequest)request).Worksheet, ((ValuesRequest)request).Range, true);
            else if (request is CellsRequest)
                return ReadExcel(workbook, ((CellsRequest)request).Worksheet, ((CellsRequest)request).Range, true);
            else
            {
                BH.Engine.Reflection.Compute.RecordError($"Requests of type {request?.GetType()} are not supported by the Excel adapter.");
                return new List<IBHoMObject>();
            }
        }


        /***************************************************/
        /**** Private Methods                           ****/
        /***************************************************/

        private List<IBHoMObject> ReadExcel(XLWorkbook workbook, string worksheet, string range, bool valuesOnly)
        {
            IXLWorksheet ixlWorksheet = Worksheet(workbook, worksheet);
            if (ixlWorksheet == null)
            {
                BH.Engine.Reflection.Compute.RecordError("No worksheets matching the request have been found.");
                return new List<IBHoMObject>();
            }

            IXLRange ixlRange = Range(ixlWorksheet, range);
            if (ixlRange == null)
            {
                Engine.Reflection.Compute.RecordError("Range provided is not in the correct format for an Excel spreadsheet.");
                return new List<IBHoMObject>();
            }

            List<DataColumn> columns = new List<DataColumn>();
            foreach (IXLRangeColumn column in ixlRange.Columns())
            {
                columns.Add(new DataColumn(column.ColumnLetter(), typeof(object)));
            }

            DataTable table = new DataTable();
            table.Columns.AddRange(columns.ToArray());

            foreach (IXLRangeRow row in ixlRange.Rows())
            {
                List<object> dataRow = new List<object>();
                foreach (IXLRangeColumn column in ixlRange.Columns())
                {
                    if (valuesOnly)
                        dataRow.Add(ixlWorksheet.Cell(row.RowNumber(), column.ColumnNumber()).GetValue<object>());
                    else
                        dataRow.Add(BH.Engine.Excel.Create.CellContents(ixlWorksheet.Cell(row.RowNumber(), column.ColumnNumber())));
                }

                table.Rows.Add(dataRow.ToArray());
            }

            return new List<IBHoMObject> { new Table { Data = table, Name = ixlWorksheet.Name } };
        }

        /***************************************************/

        private IXLWorksheet Worksheet(IXLWorkbook workbook, string worksheet)
        {
            if (!string.IsNullOrWhiteSpace(worksheet))
                return workbook.Worksheet(worksheet);
            else
                return workbook.Worksheet(0);
        }

        /***************************************************/

        private IXLRange Range(IXLWorksheet worksheet, string range)
        {
            if (!string.IsNullOrWhiteSpace(range))
                return worksheet.Range(range);
            else
                return worksheet.Range(worksheet.FirstCellUsed().Address, worksheet.LastCellUsed().Address);
        }

        /***************************************************/
    }
}
