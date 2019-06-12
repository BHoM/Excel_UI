/*
 * This file is part of the Buildings and Habitats object Model (BHoM)
 * Copyright (c) 2015 - 2018, the respective contributors. All rights reserved.
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
using BH.UI.Templates;
using BH.UI.Components;
using BH.Engine.Reflection;
using System.Linq;
using System.Collections.Generic;
using Microsoft.Office.Core;

namespace BH.UI.Excel.Components
{
    public class CreateTypeFormula : CallerFormula
    {
        /*******************************************/
        /**** Properties                        ****/
        /*******************************************/

        public override string Name {
            get {
                Type t = Caller.SelectedItem as Type;
                if (t != null)
                {

                    string ns = t.Namespace;
                    if (ns.StartsWith("BH")) ns = ns.Split('.').Skip(2).Aggregate((a, b) => $"{a}.{b}");
                    return "CreateType." + ns + "." + t.ToText();
                }
                return base.Name;
            }
        }

        public override Caller Caller { get; } = new CreateTypeCaller();

        public override string MenuRoot { get; } = "Create Type";

        /*******************************************/
        /**** Constructors                      ****/
        /*******************************************/

        public CreateTypeFormula(FormulaDataAccessor accessor, List<CommandBar> ctxMenus) : base(accessor, ctxMenus) { }

        /*******************************************/
    }
}
