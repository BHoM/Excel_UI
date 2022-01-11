/*
 * This file is part of the Buildings and Habitats object Model (BHoM)
 * Copyright (c) 2015 - 2022, the respective contributors. All rights reserved.
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
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace BH.Adapter.Excel
{
    public partial class ExcelAdapter : BHoMAdapter
    {
        /***************************************************/
        /**** Public Overrides                          ****/
        /***************************************************/

        public override List<object> Push(IEnumerable<object> objects, string tag = "", PushType pushType = PushType.AdapterDefault, ActionConfig actionConfig = null)
        {
            if (objects == null || !objects.Any())
            {
                BH.Engine.Base.Compute.RecordError("No objects were provided for Push action.");
                return new List<object>();
            }

            // If unset, set the pushType to AdapterSettings' value (base AdapterSettings default is FullCRUD).
            if (pushType == PushType.AdapterDefault)
                pushType = PushType.DeleteThenCreate;

            // Cast action config to ExcelPushConfig, create new if null.
            ExcelPushConfig config = actionConfig as ExcelPushConfig;
            if (config == null)
            {
                BH.Engine.Base.Compute.RecordNote($"{nameof(ExcelPushConfig)} has not been provided, default config is used.");
                config = new ExcelPushConfig();
            }

            // Make sure that only objects to be pushed are Tables.
            List<Type> objectTypes = objects.Select(x => x.GetType()).Distinct().ToList();
            if (objectTypes.Count != 1)
            {
                string message = "The Excel adapter only allows to push objects of a single type to a table."
                    + "\nRight now you are providing objects of the following types: "
                    + objectTypes.Select(x => x.ToString()).Aggregate((a, b) => a + ", " + b);
                Engine.Base.Compute.RecordError(message);
                return new List<object>();
            }

            Type type = objectTypes[0];
            if (type != typeof(Table))
            {
                BH.Engine.Base.Compute.RecordError($"Push failed: Excel Adapter can push only one objects of type {nameof(Table)}.");
                return new List<object>();
            }

            // Check if all tables have distinct, non-empty names.
            List<Table> tables = objects.Cast<Table>().ToList();
            if (tables.Any(x => string.IsNullOrWhiteSpace(x.Name)))
            {
                BH.Engine.Base.Compute.RecordError("Push aborted: all tables need to have non-empty name.");
                return new List<object>();
            }

            List<string> duplicateNames = tables.GroupBy(x => x.Name.ToLower()).Where(x => x.Count() != 1).Select(x => x.Key).ToList();
            if (duplicateNames.Count != 0)
            {
                BH.Engine.Base.Compute.RecordError("Push failed: all tables need to have distinct names, regardless of letter casing.\n" +
                                                        $"Following names are currently duplicate: {string.Join(", ", duplicateNames)}.");
                return new List<object>();
            }

            // Check if the workbook exists and create it if not.
            string fileName = m_FileSettings.GetFullFileName();
            XLWorkbook workbook;
            if (!File.Exists(fileName))
            {
                if (pushType == PushType.UpdateOnly)
                {
                    BH.Engine.Base.Compute.RecordError($"There is no workbook to update under {fileName}");
                    return new List<object>();
                }

                workbook = new XLWorkbook();
            }
            else
            {
                try
                {
                    workbook = new XLWorkbook(fileName);
                }
                catch (Exception e)
                {
                    BH.Engine.Base.Compute.RecordError($"The existing workbook could not be accessed due to the following error: {e.Message}");
                    return new List<object>();
                }
            }

            // Split the tables into collections to delete, create and update.
            List<Table> toDelete = new List<Table>();
            List<Table> toCreate = new List<Table>();
            List<Table> toUpdate = new List<Table>();
            switch (pushType)
            {
                case PushType.CreateNonExisting:
                    {
                        toCreate.AddRange(tables.Where(x => workbook.Worksheets.All(y => x.Name != y.Name)));
                        break;
                    }
                case PushType.DeleteThenCreate:
                    {
                        toDelete.AddRange(tables.Where(x => workbook.Worksheets.Any(y => x.Name == y.Name)));
                        toCreate.AddRange(tables);
                        break;
                    }
                case PushType.UpdateOnly:
                    {
                        toUpdate.AddRange(tables.Where(x => workbook.Worksheets.Any(y => x.Name == y.Name)));
                        break;
                    }
                case PushType.UpdateOrCreateOnly:
                    {
                        toCreate.AddRange(tables.Where(x => workbook.Worksheets.All(y => x.Name != y.Name)));
                        toUpdate.AddRange(tables.Except(toCreate).ToList());
                        break;
                    }
                default:
                    {
                        BH.Engine.Base.Compute.RecordError($"Currently Excel adapter supports only {nameof(PushType)} equal to {pushType}");
                        return new List<object>();
                    }
            }

            // Execute deletion, creation and update in a sequence.
            bool success = true;
            foreach (Table table in toDelete)
            {
                success &= Delete(workbook, table);
            }

            foreach (Table table in toCreate)
            {
                success &= Create(workbook, table, config);
            }

            foreach (Table table in toUpdate)
            {
                success &= Update(workbook, table, config);
            }

            // Try to update the workbook properties and then save it.
            try
            {
                Update(workbook, config.WorkbookProperties);
                workbook.SaveAs(fileName);
                return success ? objects.ToList() : new List<object>();
            }
            catch (Exception e)
            {
                BH.Engine.Base.Compute.RecordError($"Finalisation and saving of the workbook failed with the following error: {e.Message}");
                return new List<object>();
            }
        }

        /***************************************************/
    }
}


