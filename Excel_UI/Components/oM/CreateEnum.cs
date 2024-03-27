/*
 * This file is part of the Buildings and Habitats object Model (BHoM)
 * Copyright (c) 2015 - 2024, the respective contributors. All rights reserved.
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

using BH.Engine.Base;
using BH.UI.Base.Components;
using BH.UI.Excel.Templates;
using BH.UI.Base;
using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BH.UI.Excel.Components
{
    public class CreateEnumFormula : CallerValueListFormula
    {
        /*******************************************/
        /**** Properties                        ****/
        /*******************************************/

        public override Caller Caller { get; } = new CreateEnumCaller();

        public override string Function { get { return GetName(); } }


        /*******************************************/
        /**** Methods                           ****/
        /*******************************************/

        public override string GetName()
        {
            Type t = Caller.SelectedItem as Type;
            if (t != null)
            {
                return "CreateEnum." + t.Namespace.Split('.').Last() + "." + t.ToText();
            }
            return base.GetName();
        }

        /*******************************************/
    }
}





