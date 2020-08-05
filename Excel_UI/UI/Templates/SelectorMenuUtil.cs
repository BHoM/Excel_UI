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

using BH.oM.UI;
using BH.UI.Base;
using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BH.oM.Data;
using BH.Engine.Reflection;
using BH.UI.Base.Menus;

namespace BH.UI.Excel.Templates
{
    internal static class SelectorMenuUtil
    {
        /*******************************************/
        /**** Methods                           ****/
        /*******************************************/

        public static IExcelSelectorMenu ISetExcelSelectorMenu(IItemSelector selector)
        {
            return SetExcelSelectorMenu(selector as dynamic);
        }

        /*******************************************/

        public static IExcelSelectorMenu SetExcelSelectorMenu<T>(ItemSelector<T> selector)
        {
            var menu = new SelectorMenu_RibbonXml<T>();
            selector.SetSelectorMenu(menu);
            return menu;
        }

        /*******************************************/
    }
}

