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
using System.IO;

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

            XLWorkbook workbook;
            if (!File.Exists(fileName))
                workbook = new XLWorkbook();
            else
                workbook = new XLWorkbook(fileName);

            if (actionConfig == null)
            {
                BH.Engine.Reflection.Compute.RecordNote($"{nameof(ExcelPushConfig)} has not been provided, default config is used.");
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
                List<Table> tables = objects.Cast<Table>().ToList();
                if (tables.Any(x => string.IsNullOrWhiteSpace(x.Name)))
                {
                    BH.Engine.Reflection.Compute.RecordError("Creation aborted: all tables need to have non-empty name.");
                    return false;
                }

                List<string> duplicateNames = tables.GroupBy(x => x.Name).Where(x => x.Count() != 1).Select(x => x.Key).ToList();
                if (duplicateNames.Count != 0)
                {
                    BH.Engine.Reflection.Compute.RecordError("Creation aborted: all tables need to have distinct names." +
                                                            $"Following names are currently duplicate: {string.Join(", ", duplicateNames)}.");
                    return false;
                }

                foreach (Table table in objects.Cast<Table>())
                {
                    CreateTable(workbook, table, config);
                }
            }
            else
            {
                BH.Engine.Reflection.Compute.RecordError($"Excel Adapter can push only one objects of type {nameof(Table)}.");
                return false;
            }

            UpdateWorkbookProperties(workbook, config.WorkbookProperties);
            workbook.SaveAs(fileName);

            //TODO: why is that?
            Thread.Sleep(1000);
            return true;
        }


        /***************************************************/
        /**** Private Methods                           ****/
        /***************************************************/

        private void CreateTable(IXLWorkbook workbook, Table table, ExcelPushConfig config)
        {
            if (table?.Data == null)
            {
                BH.Engine.Reflection.Compute.RecordError("The input table is null or does not contain a table. Creation of a table aborted.");
                return;
            }

            if (string.IsNullOrWhiteSpace(table?.Name))
            {
                BH.Engine.Reflection.Compute.RecordError("The input table cannot have null name. Creation of a table aborted.");
                return;
            }

            IXLWorksheet worksheet = workbook.Worksheets.FirstOrDefault(x => x.Name == table.Name);
            if (worksheet == null)
                worksheet = workbook.AddWorksheet(table.Name);

            worksheet.Clear();

            string startingCell = config?.StartingCell;
            if (string.IsNullOrWhiteSpace(startingCell))
                startingCell = "A1";

            try
            {
                worksheet.Cell(startingCell).InsertData(table.Data);
            }
            catch(Exception e)
            {
                BH.Engine.Reflection.Compute.RecordError($"Population of worksheet {table.Name} failed with the following error: {e.Message}");
            }
        }

        /***************************************************/
    }
}

