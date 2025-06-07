/*
 * This file is part of the Buildings and Habitats object Model (BHoM)
 * Copyright (c) 2015 - 2025, the respective contributors. All rights reserved.
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

using BH.Adapter;
using BH.oM.Adapter;
using BH.oM.Base;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using BH.UI.Excel;


namespace BH.UI.Excel
{
    public partial class AddIn : IExcelAddIn
    {
        /*******************************************/
        /**** Public Methods                    ****/
        /*******************************************/
        public static void SetAdapter(Range selection)
        {
            if(selection.Count != 1)
            {
                BH.Engine.Base.Compute.RecordError("Only one Adapter is accepted !");
                return;
            }

            object value = selection.Value;

            if (value == null)
            {
                m_Adapter = null;
                return;
            }

            object obj = GetObject(value as string);

            if (obj == null)
            {
                m_Adapter = null;
                return;
            }

            BHoMAdapter adapter = obj as BHoMAdapter;

            if (adapter != null) 
            {
                m_Adapter = adapter;
                m_AdapterName = selection.Value as string;
            }
            else
            {
                m_Adapter = null;
            }
        }

        /*******************************************/

        public static string GetAdapterName()
        {
            if (m_Adapter != null)
            {
                return m_AdapterName;
            }
            else
            {
                return "----";
            }
        }

        /*******************************************/

        public static void Execute(string command, Range objects)
        {
            Type commandType = BH.Engine.Base.Create.Type($"BH.oM.Adapter.Commands.{command}");
            dynamic runCommand = Activator.CreateInstance(commandType);

            List<IBHoMObject> target = new List<IBHoMObject>();
            foreach (Range cell in objects)
            {
                object value = cell.Value;
                if (value != null)
                {
                    // Store the item if exists
                    string id = GetId(cell.Value as string);
                    object item = GetObject(id);
                    target.Add(item as IBHoMObject);
                }
            }

            if (target.Count == 0)
            {
                return;
            }

            commandType.GetProperty("Identifiers")?.SetValue(runCommand, target.Cast<IObject>().ToList());

            m_Adapter.Execute(runCommand,actionConfig : null);

        }

        /*******************************************/

        public static string Execute(string command)
        {
            Type commandType = BH.Engine.Base.Create.Type($"BH.oM.Adapter.Commands.{command}");
            dynamic runCommand = Activator.CreateInstance(commandType);
            var output = m_Adapter.Execute(runCommand, actionConfig: null);

            if (output == null || output.Item1 == null|| output.Item1.Count == 0)
            {
                return null;
            }
            
            object result = ToExcel(output.Item1);

            if (output.Item1 != null)
            {                
                string id = GetId(result as string);
                m_InternalisedData[id] = output.Item1;
                WriteJsonToSheet("BHoM_DataHidden", m_InternalisedData);
            }

            return (string) result;
        }

        /*******************************************/

        public static void ExecuteCustomCommand(string command, Range objects)
        {
            BH.oM.Adapter.Commands.CustomCommand customCommand = new oM.Adapter.Commands.CustomCommand();
            customCommand.Command = command;
            List<IBHoMObject> target = new List<IBHoMObject>();

            foreach (Range cell in objects)
            {
                object value = cell.Value;
                if (value != null)
                {
                    // Store the item if exists
                    string id = GetId(cell.Value as string);
                    object item = GetObject(id);
                    target.Add(item as IBHoMObject);
                }
            }

            if (target.Count == 0)
            {
                return;
            }

            customCommand.Parameters = new Dictionary<string, object>()
            {
                { "Parameters" ,target },
            };

            m_Adapter.Execute(customCommand);
        }

        /*******************************************/


        /*******************************************/
        /**** Private Fields                   *****/
        /*******************************************/
        private static BHoMAdapter m_Adapter;
        private static string m_AdapterName = "----";
    }
}






