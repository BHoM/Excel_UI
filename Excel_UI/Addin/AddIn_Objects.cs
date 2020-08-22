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

using System;
using System.IO;
using System.Reflection;
using System.Linq;
using ExcelDna.Integration;
using System.Collections.Generic;
using System.Collections;
using BH.Engine.Reflection;
using BH.oM.Base;
using System.Linq.Expressions;
using BH.UI.Base;
using BH.UI.Excel.Templates;
using BH.UI.Excel.Components;
using BH.UI.Excel.Global;
using BH.UI.Base.Global;
using BH.UI.Base.Components;
using System.Runtime.InteropServices;
using NetOffice.ExcelApi;
using NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;
using NetOffice.ExcelApi.Enums;
using System.Drawing;
using System.Xml;
using BH.oM.UI;
using BH.Engine.Base;

namespace BH.UI.Excel
{
    public partial class AddIn : IExcelAddIn
    {
        /*******************************************/
        /**** Methods                           ****/
        /*******************************************/

        public static string IAddObject(object item)
        {
            return AddObject(item as dynamic);
        }

        /*******************************************/

        public static void IAddObject(object item, string id)
        {
            m_Objects[id] = item;
        }

        /*******************************************/

        public static object GetObject(string key)
        {
            // Make sure the key is an id
            string id = GetId(key);
            if (id.Length > 0)
                key = id;

            // Return the object if in dictionary, return null otherwise
            if (m_Objects.ContainsKey(key))
                return m_Objects[key];
            else
                return null;
        }

        /*******************************************/

        public static string GetId(string key)
        {
            // Make sure the key is an id
            int start = key.LastIndexOf("[");
            int end = key.LastIndexOf("]");
            if (start != -1 && end != -1 && end > start)
                return key.Substring(++start, end - start);
            else
                return "";
        }

        /*******************************************/

        public static void ClearObjects()
        {
            m_Objects.Clear();
        }

        /*******************************************/
        /**** Private Methods                   ****/
        /*******************************************/

        private static string AddObject(object item)
        {
            string id = ToString(Guid.NewGuid());
            m_Objects[id] = item;

            return id;
        }

        /*******************************************/

        private static string AddObject(IBHoMObject item)
        {
            string id = ToString(item.BHoM_Guid);
            m_Objects[id] = item;

            return id;
        }

        /*****************************************/

        private static string ToString(Guid id)
        {
            return System.Convert.ToBase64String(id.ToByteArray()).Remove(8);
        }


        /*******************************************/
        /**** Private Fields                    ****/
        /*******************************************/

        private static Dictionary<string, object> m_Objects = new Dictionary<string, object>(); //TODO: This grows very quickly -> need to find a way to remove old objects too

        /*******************************************/
    }
}

