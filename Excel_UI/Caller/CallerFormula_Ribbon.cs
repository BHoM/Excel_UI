/*
 * This file is part of the Buildings and Habitats object Model (BHoM)
 * Copyright (c) 2015 - 2024, the respective contributors. All rights reserved.
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

using BH.UI.Base;
using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Xml;
using System.Collections;

namespace BH.UI.Excel.Templates
{
    public abstract partial class CallerFormula
    {
        /*******************************************/
        /**** Methods                           ****/
        /*******************************************/

        public virtual string GetRibbonXml()
        {
            XmlDocument doc = new XmlDocument();
            XmlElement element;
            
            if (Caller.HasPossibleItems) // Multi-Choice caller -> drop down menu
            {
                element = doc.CreateElement("dynamicMenu");
                element.SetAttribute("getContent", "GetContent");
            }
            else // Single-Choice caller -> button
            {
                element = doc.CreateElement("button");
                element.SetAttribute("tag", Caller.GetType().Name);
                element.SetAttribute("onAction", "FillFormula");
            }
            
            // Configure the menu element
            element.SetAttribute("id", Caller.GetType().Name);
            element.SetAttribute("getImage", "GetImage");
            element.SetAttribute("label", Caller.Name);
            element.SetAttribute("screentip", Caller.Name);

            if (Caller.Description != "")
                element.SetAttribute("supertip", Caller.Description);

            return element.OuterXml;
        }

        /*******************************************/

        public virtual string GetInnerRibbonXml()
        {
            m_Menu = new SelectorMenu_RibbonXml();
            m_Menu.RootName = Caller.GetType().Name;
            Caller.SetSelectorMenu(m_Menu);
            
            XmlDocument doc = new XmlDocument();
            XmlElement root = doc.CreateElement("root");
            Caller.AddToMenu(root);

            XmlElement menu = root.FirstChild as XmlElement;
            if (menu == null)
                return "";
            menu.RemoveAllAttributes();
            menu.SetAttribute("xmlns", "http://schemas.microsoft.com/office/2006/01/customui");
            return root.InnerXml;
        }

        /*******************************************/

        public virtual void Select(string id)
        {
            CallerFormula formula = this;
            if (m_Menu != null)
            {
                CallerFormula newInstance = AddIn.InstantiateCaller(Caller.GetType().Name, m_Menu.GetItem(id));
                if (newInstance != null)
                    formula = newInstance;
            }

            ExcelAsyncUtil.QueueAsMacro(() => formula.FillFormula(AddIn.CurrentSelection())); 
        }


        /*******************************************/
        /**** Private Fields                    ****/
        /*******************************************/

        private SelectorMenu_RibbonXml m_Menu;

        /*******************************************/
    }
}





