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
using BH.oM.Base;
using BH.UI.Excel.Templates;
using BH.UI.Base;
using Microsoft.Office.Core;
using System.Collections.Generic;
using BH.UI.Excel.Callers;
using System.Linq;

namespace BH.UI.Excel.Components
{
    // Bespoke Excel explode method
    public class ExplodeFormula : CallerFormula
    {
        /*******************************************/
        /**** Properties                        ****/
        /*******************************************/

        public override Caller Caller { get; } = new ExplodeCaller();

        public override string Function { get; } = "BHoM.Explode";

        public override string Category { get; } = "Engine";


        /*******************************************/
        /**** Override Methods                  ****/
        /*******************************************/

        public override object Run(object[] inputs)
        {
            base.Run(inputs);
            return AddIn.ToExcel(m_DataAccessor.Outputs[0] as List<List<object>>);

        }


        /*******************************************/
    }
}






