/*
 * This file is part of the Buildings and Habitats object Model (BHoM)
 * Copyright (c) 2015 - 2021, the respective contributors. All rights reserved.
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
            XLWorkbook workbook = null;
            try
            {
                workbook = new XLWorkbook(m_FileSettings.GetFullFileName());
            }
            catch
            {
                // No error raised here because it will get raised under if (workbook == null) below.
            }

            if (workbook == null)
            {
                BH.Engine.Reflection.Compute.RecordError("The file under location specified in the settings is not a valid Excel workbook.");
                return new List<IBHoMObject>();
            }

            if (request is CellValuesRequest)
                return ReadExcel(workbook, ((CellValuesRequest)request).Worksheet, ((CellValuesRequest)request).Range, true);
            else if (request is CellContentsRequest)
                return ReadExcel(workbook, ((CellContentsRequest)request).Worksheet, ((CellContentsRequest)request).Range, false);
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
            List<IBHoMObject> result = new List<IBHoMObject>();
            foreach (IXLWorksheet ixlWorksheet in Worksheets(workbook, worksheet))
            {
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

                result.Add(new Table { Data = table, Name = ixlWorksheet.Name });
            }

            return result;
        }

        /***************************************************/

        private IEnumerable<IXLWorksheet> Worksheets(IXLWorkbook workbook, string worksheet)
        {
            if (!string.IsNullOrWhiteSpace(worksheet))
            {
                try
                {
                    return new List<IXLWorksheet> { workbook.Worksheet(worksheet) };
                }
                catch
                {
                    BH.Engine.Reflection.Compute.RecordError("No worksheets matching the request have been found.");
                    return new List<IXLWorksheet>();
                }
            }
            else
                return workbook.Worksheets;
        }

        /***************************************************/

        private IXLRange Range(IXLWorksheet worksheet, string range)
        {
            if (!string.IsNullOrWhiteSpace(range))
            {
                try
                {
                    return worksheet.Range(range);
                }
                catch
                {
                    return null;
                }
            }
            else
                return worksheet.Range(worksheet.FirstCellUsed().Address, worksheet.LastCellUsed().Address);
        }

        /***************************************************/
    }
}

