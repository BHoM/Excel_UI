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

using BH.oM.Base;
using BH.UI.Components;
using BH.UI.Excel.Templates;
using BH.UI.Templates;
using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BH.UI.Excel.Components
{
    public class CreateObjectFormula : CallerFormula
    {
        /*******************************************/
        /**** Properties                        ****/
        /*******************************************/

        public override Caller Caller { get; } = new CreateObjectCaller();

        public override string MenuRoot { get; } = "Create Object";

        public override string Name
        {
            get
            {
                if (Caller is MethodCaller && Caller.SelectedItem != null)
                {
                    Type decltype = (Caller as MethodCaller).OutputParams.First().DataType;
                    if (decltype.IsSubclassOf(typeof(IObject)))
                    {
                        string ns = decltype.Namespace;
                        if (ns.StartsWith("BH")) ns = ns.Split('.').Skip(2).Aggregate((a, b) => $"{a}.{b}");
                        return "Create." + ns + "." + Caller.Name;
                    }
                    return base.Name;

                }
                return Category + "." + Caller.Name;
            }
        }

        /*******************************************/
        /**** Constructors                      ****/
        /*******************************************/

        public CreateObjectFormula(FormulaDataAccessor accessor) : base(accessor) { }
    }
}
