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

using BH.UI.Excel.Templates;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using NetOffice.ExcelApi;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace BH.UI.Excel.Addin
{
    public partial class Ribbon : ExcelRibbon
    {
        /*******************************************/
        /**** Methods                           ****/
        /*******************************************/

        public void Internalise(IRibbonControl control)
        {
            Application app = Application.GetActiveInstance();
            Range selected = app.Selection as Range;

            foreach (Range objcell in selected)
            {
                string value;
                try
                {
                    value = (string)objcell.Value;
                    if (value == null || value.Length == 0) continue;
                }
                catch { continue; }

                Project proj = Project.ForIDs(new string[] { value });

                if (proj.Count((o) => !(o is Adapter.BHoMAdapter)) == 0) continue;
                proj.SaveData(app.ActiveWorkbook);

                objcell.Value = value;
            }
        }

        /*******************************************/
    }
}
