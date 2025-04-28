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
using Microsoft.Office.Interop.Excel;
using BH.Engine.Serialiser;
using BH.oM.Base;
using BH.UI.Excel.Templates;
using BH.UI.Base.Components;

namespace BH.UI.Excel
{
    public partial class AddIn : IExcelAddIn
    {
        /*******************************************/
        /**** Methods                           ****/
        /*******************************************/

        public static void Internalise(Range selection)
        {
            // Get the FromJson caller 
            string callerName = typeof(FromJsonCaller).Name;
            if (!CallerShells.ContainsKey(callerName))
                return;
            CallerFormula formula = CallerShells[callerName];

            // Make sure it is registered
            Register(formula);

            // Replace the function to recover the internalised data
            foreach (Range cell in selection)
            {
                object value = cell.Value; 
                if (value != null)
                {
                    // Store the item if exists
                    string id = GetId(cell.Value as string);
                    object item = GetObject(id);
                    if (item != null)
                        m_InternalisedData[id] = item;

                    // Replace cell formula with value
                    ExcelAsyncUtil.QueueAsMacro(() => { cell.Formula = value; });
                }
            }
            
            // Save it to the hidden sheet
            WriteJsonToSheet("BHoM_DataHidden", m_InternalisedData); //TODO: if we set the formula above to a "FromJson" call, we can get rid of the hidden sheet
        }

        /*******************************************/

        public static void RestoreData()
        {
            // Make sure the data in the hidden sheet has been loaded to the dictionary
            if (m_InternalisedData.Count == 0)
                m_InternalisedData = ReadJsonFromSheet("BHoM_DataHidden");

            // Update cells of active sheets based on internalised data
            foreach (var kvp in m_InternalisedData)
                IAddObject(kvp.Value, kvp.Key);
        }


        /*******************************************/
        /**** Private Methods                   ****/
        /*******************************************/

        private static void WriteJsonToSheet(string sheetName, Dictionary<string, object> dic)
        {
            // Get the data sheet
            Worksheet sheet = Sheet(sheetName, true, true);
            if (sheet == null)
                return;

            // Save the dictionary as json
            int index = 1;
            const int characterLimit = 32000; // The real number is 32,767 but let's keep it tidy
            string json = dic.ToJson();
            for (int i = 0; i < json.Length; i += characterLimit)
            {
                try
                {
                    sheet.Cells[index, 1].Value = json.Substring(i, Math.Min(characterLimit, json.Length - i));
                    index++;
                }
                catch { }
            } 
        }

        /*******************************************/

        private static Dictionary<string, object> ReadJsonFromSheet(string sheetName)
        {
            // Get the hidden worksheet
            Worksheet sheet = Sheet(sheetName, false);
            if (sheet == null)
                return new Dictionary<string, object>();

            // Get the json version of the internalised data
            string json = "";
            for (int i = 1; i < 1000; i++) // Just for safety
            {
                try
                {
                    string segment = sheet.Cells[i, 1].Value as string;
                    if (string.IsNullOrEmpty(segment))
                        break;
                    else
                        json += segment;
                }
                catch
                {
                    break;
                }
            }

            // Extrat dictionary from json
            try
            {
                Dictionary<string, object> custom = Engine.Serialiser.Convert.FromJson(json) as Dictionary<string, object>; // This is because the serialiser engine deserilaise top Dictionary<string, object> as CostomObject at the moment
                if (custom != null)
                    return custom;
            }
            catch { }

            // Return empty dictionary if not successful
            return new Dictionary<string, object>();
        }


        /*******************************************/
        /**** Private Fields                    ****/
        /*******************************************/

        private static Dictionary<string, object> m_InternalisedData = new Dictionary<string, object>();

        /*******************************************/
    }
}






