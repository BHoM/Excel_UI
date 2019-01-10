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

namespace BH.UI.Excel.Templates
{
    public abstract class SingleOptionCallerFormula : CallerFormula
    {
        /*******************************************/
        /**** Constructors                      ****/
        /*******************************************/

        public SingleOptionCallerFormula(FormulaDataAccessor accessor) : base(accessor)
        {
            var commandBar = Application.CommandBars["Cell"];

            var menu = commandBar.FindControl(
                Type: MsoControlType.msoControlPopup,
                Tag: Category
                ) as CommandBarPopup;

            if (menu == null)
            {
                menu = commandBar.Controls.Add(MsoControlType.msoControlPopup, Temporary: true) as CommandBarPopup;
                menu.Caption = Category;
                menu.Tag = Category;
            }

            m_btn = menu.Controls.Add(MsoControlType.msoControlButton, Temporary: true) as CommandBarButton;
            m_btn.Tag = MenuRoot + "_btn";
            m_btn.Caption = MenuRoot;
            m_btn.Click += OnClick;
        }

        /*******************************************/
        /**** Private Methods                   ****/
        /*******************************************/

        private void OnClick(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            Application.SendKeys("="+Function+"{(}",true);
            if(Caller.InputParams.Count == 0)
            {
                Application.SendKeys("{)}{ENTER}",true);
            }
        }

        /*******************************************/
        /**** Private Fields                    ****/
        /*******************************************/

        private CommandBarButton m_btn;
    }
}
