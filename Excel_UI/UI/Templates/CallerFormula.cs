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
using System.Text;
using System.Threading.Tasks;

namespace BH.UI.Excel.Templates
{
    public abstract class CallerFormula
    {
        /*******************************************/
        /**** Properties                        ****/
        /*******************************************/

        public virtual string Name 
        {
            get
            {
                if (Caller is MethodCaller && Caller.SelectedItem != null)
                {
                    Type decltype = (Caller as MethodCaller).Method.DeclaringType;
                    return decltype.Name + "." + decltype.Namespace.Split('.').Last() + "." + Caller.Name;
                }
                return Category + "." + Caller.Name;
            }
        }

        public virtual string Category { get { return Caller.Category; } }

        public abstract string MenuRoot { get; }

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
                        .Select(p => p.DataType.ToText())
                        .Select(p => p.Replace("[]", "s"))
                        .Select(p => p.Replace("[,]", "Matrix"))
                        .Select(p => p.Replace("&", ""))
                        .Select(p => p.Replace("<", "Of"))
                        .Select(p => p.Replace(">", ""))
                        .Select(p => p.Replace(", ", "_"))
                        .Select(p => p.Replace("`", "_"))
                        .Aggregate((a, b) => $"{a}_{b}");
                }

                return Name + params_;
            }
        }

        public abstract Caller Caller { get; }

        /*******************************************/
        /**** Constructors                      ****/
        /*******************************************/

        public CallerFormula(FormulaDataAccessor accessor)
        {
            m_dataAccessor = accessor;
            Caller.SetDataAccessor(m_dataAccessor);

            if (Caller.Selector != null)
            {
                var smenu = SelectorMenuUtil.ISetExcelSelectorMenu(Caller.Selector);
                smenu.RootName = MenuRoot;
            }

            Application = ExcelDna.Integration.ExcelDnaUtil.Application as Application;

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

            Caller.AddToMenu(menu.Controls);
            Caller.ItemSelected += Caller_ItemSelected;
        }

        /*******************************************/
        /**** Private Methods                   ****/
        /*******************************************/

        private void Caller_ItemSelected(object sender, object e)
        {
            Application.SendKeys("="+Function+"{(}",true);
            if(Caller.InputParams.Count == 0)
            {
                Application.SendKeys("{)}{ENTER}",true);
            }
        }

        /*******************************************/
        /**** Private Properties                ****/
        /*******************************************/

        protected Application Application { get; private set; }

        /*******************************************/
        /**** Private Fields                    ****/
        /*******************************************/

        private FormulaDataAccessor m_dataAccessor;
    }
}
