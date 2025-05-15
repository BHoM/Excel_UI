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

using System;
using System.IO;
using System.Reflection;
using System.Linq;
using ExcelDna.Integration;
using System.Collections.Generic;
using System.Collections;
using System.Linq.Expressions;
using BH.UI.Excel.Templates;
using BH.oM.Adapter;
using BH.Adapter;
using BH.oM.Base;
using Microsoft.Office.Interop.Excel;


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

        public static void Execute(IExecuteCommand command, ActionConfig actionConfig = null)
        {
            m_Adapter.Execute(command, actionConfig);
        }

        /*******************************************/

        public static void Execute(string command, Range objects )
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
                { "BHoMObjects" ,target }
            };

            m_Adapter.Execute(customCommand);
        }

        /*******************************************/
        /**** Private Fields                   *****/
        /*******************************************/
        private static BHoMAdapter m_Adapter;
        private static string m_AdapterName = "----";
    }
}






