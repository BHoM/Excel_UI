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
        /**** Properties                        ****/
        /*******************************************/

        public static Dictionary<string, CallerFormula> Callers { get; private set; }


        /*******************************************/
        /**** Constructors                      ****/
        /*******************************************/

        static AddIn()
        {
            // Collect the callers from assemblies
            Callers = ExcelIntegration.GetExportedAssemblies()
                .SelectMany(a => a.GetTypes())
                .Where(t => t.Namespace == "BH.UI.Excel.Components" && typeof(CallerFormula).IsAssignableFrom(t))
                .Select(t => t.GetConstructor(new Type[] { }).Invoke(new object[] { }) as CallerFormula)
                .ToDictionary(o => o.Caller.GetType().Name);

            // Add the event listeners
            foreach (var formula in Callers.Values)
            {
                formula.OnRun += (s, e) =>
                {
                    var f = (s as CallerFormula);
                    var caller = f.Caller;
                    string name = Engine.Excel.Query.Filename();
                    var manager = ComponentManager.GetManager(name);
                    if (manager != null)
                    {
                        manager.Store(caller, f.Function);
                    }
                };
            }
        }

        /*******************************************/
    }
}

