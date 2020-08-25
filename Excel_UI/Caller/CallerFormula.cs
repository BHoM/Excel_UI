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
        /**** Properties                        ****/
        /*******************************************/

        public virtual string Category { get { return Caller.Category; } }

        public abstract Caller Caller { get; }

        public virtual string Function
        {
            get
            {

                IEnumerable<ParamInfo> paramList = Caller.InputParams;
                bool hasParams = paramList.Count() > 0;
                string params_ = "";
                if (hasParams)
                {
                    params_ = "?by_" + paramList
                        .Select(p => p.DataType.ToText(false, false, "Of", "_", ""))
                        .Select(p => p.Replace("[]", "s"))
                        .Select(p => p.Replace("[,]", "Matrix"))
                        .Select(p => p.Replace("&", ""))
                        .Select(p => p.Replace("`", "_"))
                        .Aggregate((a, b) => $"{a}_{b}");
                }

                return GetName() + params_;
            }
        }


        /*******************************************/
        /**** Constructors                      ****/
        /*******************************************/

        public CallerFormula()
        {
            m_DataAccessor = new FormulaDataAccessor();
            Caller.SetDataAccessor(m_DataAccessor);
        }

        /*******************************************/
        /**** Methods                           ****/
        /*******************************************/

        public virtual string GetName()
        {
            if (Caller.SelectedItem != null && Caller.SelectedItem is MethodBase)
            {
                Type decltype = ((MethodBase)Caller.SelectedItem).DeclaringType;
                string ns = decltype.Namespace;
                if (ns.StartsWith("BH"))
                    ns = ns.Split('.').Skip(2).Aggregate((a, b) => $"{a}.{b}");
                return decltype.Name + "." + ns + "." + Caller.Name;
            }
            return Category + "." + Caller.Name;
        }


        /*******************************************/
        /**** Private Fields                    ****/
        /*******************************************/

        protected FormulaDataAccessor m_DataAccessor = null;

        /*******************************************/
    }
}

