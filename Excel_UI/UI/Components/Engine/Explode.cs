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

namespace BH.UI.Excel.Components
{
    public class ExplodeFormula : SingleOptionCallerFormula
    {
        /*******************************************/
        /**** Properties                        ****/
        /*******************************************/

        // Bespoke Excel explode method
        public override Caller Caller { get; } = new MethodCaller(typeof(Properties).GetMethod("Explode"));

        public override string Function { get; } = "BHoM.Explode";
            
        public override string Category { get; } = "Engine";

        /*******************************************/
        /**** Constructors                      ****/
        /*******************************************/

        public ExplodeFormula(FormulaDataAccessor accessor) : base(accessor) { }

        /*******************************************/
    }
}
