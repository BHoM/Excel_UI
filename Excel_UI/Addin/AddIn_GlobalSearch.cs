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

using System;
using System.IO;
using System.Reflection;
using System.Linq;
using ExcelDna.Integration;
using System.Collections.Generic;
using System.Collections;
using BH.UI.Excel.Templates;
using BH.UI.Base.Global;


namespace BH.UI.Excel
{
    public partial class AddIn : IExcelAddIn
    {
        /*******************************************/
        /**** Methods                           ****/
        /*******************************************/

        [ExcelCommand(ShortCut = "^B")]
        public static void OpenGlobalSearch()
        {
            m_CurrentSelection = CurrentSelection();
            var control = new System.Windows.Forms.ContainerControl();
            m_GlobalSearch.SetParent(control);
        }


        /*******************************************/
        /**** Private Methods                   ****/
        /*******************************************/

        protected void InitGlobalSearch()
        {
            if (m_GlobalSearch == null)
            {
                try
                {
                    m_GlobalSearch = new SearchMenu_WinForm();
                    m_GlobalSearch.ItemSelected += GlobalSearch_ItemSelected;
                }
                catch (Exception e)
                {
                    Engine.Base.Compute.RecordError(e.Message);
                }
            }
        }

        /*******************************************/

        protected void GlobalSearch_ItemSelected(object sender, oM.UI.ComponentRequest e)
        {
            if (e != null && e.CallerType != null)
            {
                CallerFormula formula = InstantiateCaller(e.CallerType.Name, e.SelectedItem);
                if (formula != null)
                    formula.FillFormula(m_CurrentSelection);
            }
        }

        /*******************************************/
        /**** Private Fields                    ****/
        /*******************************************/

        private static SearchMenu m_GlobalSearch = null;
        private static ExcelReference m_CurrentSelection = null;

        /*******************************************/
    }
}





