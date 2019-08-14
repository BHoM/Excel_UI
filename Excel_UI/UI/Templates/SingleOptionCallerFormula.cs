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

using BH.Engine.Reflection;
using BH.oM.UI;
using BH.UI.Templates;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace BH.UI.Excel.Templates
{
    public abstract class SingleOptionCallerFormula : CallerFormula
    {
        /*******************************************/
        /**** Constructors                      ****/
        /*******************************************/

        public SingleOptionCallerFormula() : base()
        {
        }

        public override string GetRibbonXml()
        {
            XmlDocument doc = new XmlDocument();
            XmlElement btn = doc.CreateElement("button");
            btn.SetAttribute("id", Caller.GetType().Name);
            btn.SetAttribute("tag", Caller.GetType().Name);
            btn.SetAttribute("getImage", "GetImage");
            btn.SetAttribute("label", Caller.Name);
            btn.SetAttribute("screentip", Caller.Name);
            btn.SetAttribute("supertip", Caller.Description);
            btn.SetAttribute("onAction","FillFormula");
            return btn.OuterXml;
        }

        /*******************************************/

        public override void Select(string id)
        {
            FillFormula();
        }

        /*******************************************/
        /**** Private Fields                    ****/
        /*******************************************/

        public override string MenuRoot => "";
    }
}
