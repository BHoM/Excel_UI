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

using BH.Engine.Reflection;
using BH.Engine.Excel;
using BH.oM.UI;
using BH.UI.Base;
using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using NetOffice.ExcelApi;
using System.Xml;
using BH.Engine.Serialiser;
using System.Reflection;
using System.Linq.Expressions;
using System.Collections;

namespace BH.UI.Excel.Templates
{
    public abstract partial class CallerFormula
    {
        /*******************************************/
        /**** Methods                           ****/
        /*******************************************/

        public virtual object Run(object[] inputs)
        {
            //Clear current events
            Engine.Reflection.Compute.ClearCurrentEvents();

            // Run the caller
            m_DataAccessor.SetInputs(inputs.ToList(), Caller.InputParams.Select(x => x.DefaultValue).ToList());
            Caller.Run();
            object result = m_DataAccessor.GetOutputs();

            // Handle possible errors
            var errors = Engine.Reflection.Query.CurrentEvents().Where(e => e.Type == oM.Reflection.Debugging.EventType.Error);
            if (errors.Count() > 0)
                Engine.Excel.Query.Caller().Note(errors.Select(e => e.Message).Aggregate((a, b) => a + "\n" + b));
            else
                Engine.Excel.Query.Caller().Note("");

            // Return result
            return result;
        }

        /*******************************************/
    }
}

