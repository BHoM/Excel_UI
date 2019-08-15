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

using BH.oM.UI;
using BH.UI.Templates;
using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BH.Engine.Reflection;
using BH.oM.Data.Collections;
using NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;
using System.Xml;

namespace BH.UI.Excel.Templates
{
    public class SelectorMenu_RibbonXml<T> : SelectorMenu<T, XmlElement>, IExcelSelectorMenu
    {
        /*******************************************/
        /**** Properties                        ****/
        /*******************************************/

        public string RootName { get; set; }

        /*******************************************/
        /**** Constructors                      ****/
        /*******************************************/

        public SelectorMenu_RibbonXml() : base(null, null)
        {
        }

        /*******************************************/
        /**** Protected Methods                 ****/
        /*******************************************/

        protected override void AddSearchBox(XmlElement menu, List<SearchItem> itemList)
        {
            // Noop
        }

        /*******************************************/

        protected override void AddTree(XmlElement menu, Tree<T> itemTree)
        {
            AppendMenuTree(itemTree, menu);
        }

        /*******************************************/
        /**** Private Methods                   ****/
        /*******************************************/

        private void AppendMenuTree(Tree<T> tree, XmlElement menu)
        {
            XmlDocument document = menu.OwnerDocument;
            XmlElement element;
            string id = "id"+Guid.NewGuid().ToString();
            if (tree.Children.Count > 0)
            {
                element = document.CreateElement("menu");
                foreach (Tree<T> childTree in tree.Children.Values.OrderBy(x => x.Name))
                    AppendMenuTree(childTree, element);
            }
            else
            {
                T method = tree.Value;
                element = document.CreateElement("button");
                element.SetAttribute("onAction", "FillFormula");
                element.SetAttribute("supertip", method.IDescription());
                m_ItemLinks[id] = method;
            }
            element.SetAttribute("label", tree.Name);
            element.SetAttribute("id", id);
            element.SetAttribute("tag", RootName);
            menu.AppendChild(element);
        }

        public void Select(string id)
        {
            if(m_ItemLinks.ContainsKey(id))
                ReturnSelectedItem(m_ItemLinks[id]);
        }

        /*******************************************/
        /**** Private Fields                    ****/
        /*******************************************/

        private Dictionary<string, T> m_ItemLinks = new Dictionary<string, T>();
    }
}
