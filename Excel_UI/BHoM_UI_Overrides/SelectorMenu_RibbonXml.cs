/*
 * This file is part of the Buildings and Habitats object Model (BHoM)
 * Copyright (c) 2015 - 2025, the respective contributors. All rights reserved.
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
using BH.Engine.Reflection;
using BH.oM.Data.Collections;
using System.Xml;
using BH.UI.Base.Menus;

namespace BH.UI.Excel.Templates
{
    public class SelectorMenu_RibbonXml : ItemSelectorMenu<XmlElement>
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
        /**** Methods                           ****/
        /*******************************************/

        public void Select(string id)
        {
            if(m_ItemLinks.ContainsKey(id))
                ReturnSelectedItem(m_ItemLinks[id]);
        }

        /*******************************************/

        public object GetItem(string id)
        {
            if (m_ItemLinks.ContainsKey(id))
                return m_ItemLinks[id];
            else
                return null;
        }


        /*******************************************/
        /**** Private Methods                   ****/
        /*******************************************/

        protected override void AddSearchBox(XmlElement menu, List<SearchItem> itemList)
        {
            // Noop
        }

        /*******************************************/

        protected override void AddTree(XmlElement menu, Tree<object> itemTree)
        {
            AppendMenuTree(itemTree, menu);
        }

        /*******************************************/

        private void AppendMenuTree(Tree<object> tree, XmlElement menu)
        {
            XmlDocument document = menu.OwnerDocument;
            XmlElement element;
            string id = "id"+Guid.NewGuid().ToString();
            if (tree.Children.Count > 0)
            {
                element = document.CreateElement("menu");
                foreach (Tree<object> childTree in tree.Children.Values.OrderBy(x => x.Name))
                    AppendMenuTree(childTree, element);
            }
            else
            {
                object method = tree.Value;
                element = document.CreateElement("button");
                element.SetAttribute("onAction", "FillFormula");
                string description = method.IDescription();
                if(description.Length > 0)
                {
                    // Ribbon XML schema has a hard limit of 1024 characters, truncate if we exceed it
                    if (description.Length > 1024)
                        description = description.Substring(0, 1024);
                    element.SetAttribute("supertip", description);
                }
                m_ItemLinks[id] = method;
            }
            element.SetAttribute("label", tree.Name);
            element.SetAttribute("id", id);
            element.SetAttribute("tag", RootName);
            menu.AppendChild(element);
        }

        /*******************************************/
        /**** Private Fields                    ****/
        /*******************************************/

        private Dictionary<string, object> m_ItemLinks = new Dictionary<string, object>();

        /*******************************************/
    }
}






