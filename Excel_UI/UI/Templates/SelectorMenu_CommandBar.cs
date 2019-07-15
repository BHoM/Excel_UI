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

namespace BH.UI.Excel.Templates
{
    public class SelectorMenu_CommandBar<T> : SelectorMenu<T, CommandBarControls>, IExcelSelectorMenu
    {
        /*******************************************/
        /**** Properties                        ****/
        /*******************************************/

        public string RootName { get; set; }

        /*******************************************/
        /**** Constructors                      ****/
        /*******************************************/

        public SelectorMenu_CommandBar() : base(null, null) { }

        /*******************************************/
        /**** Protected Methods                 ****/
        /*******************************************/

        protected override void AddSearchBox(CommandBarControls menu, List<SearchItem> itemList)
        {
            // Noop
        }

        /*******************************************/

        protected override void AddTree(CommandBarControls menu, Tree<T> itemTree)
        {
            itemTree.Name = RootName;
            AppendMenuTree(itemTree, menu);
        }

        /*******************************************/
        /**** Private Methods                   ****/
        /*******************************************/

        private void AppendMenuTree(Tree<T> tree, CommandBarControls menu)
        {
            try
            {
                if (tree.Children.Count > 0)
                {
                    CommandBarControls treeMenu = AppendMenuItem(menu, tree.Name);
                    foreach (Tree<T> childTree in tree.Children.Values.OrderBy(x => x.Name))
                        AppendMenuTree(childTree, treeMenu);
                }
                else
                {
                    T method = tree.Value;
                    CommandBarButton methodItem = AppendMenuItem(menu, tree.Name, Item_Click);
                    methodItem.Tag = Guid.NewGuid().ToString();
                    m_ItemLinks[methodItem.Tag] = method;
                }
            }
            catch (Exception e)
            {
                Compute.RecordError(e.Message);
            }
        }

        /*******************************************/

        private void Item_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                T method = m_ItemLinks[Ctrl.Tag];
                ReturnSelectedItem(method);
            }
            catch { }
        }

        /*******************************************/

        private CommandBarButton AppendMenuItem(CommandBarControls menu, string name, CommandBarButton_ClickEventHandler onClick)
        {
            CommandBarButton btn = menu.Add(MsoControlType.msoControlButton, null, null, null, true) as CommandBarButton;
            btn.Caption = name;
            btn.ClickEvent += onClick;
            // Otherwise it's GC'd and the click handler isn't run
            m_buttons.Add(btn); 
            return btn;
        }

        /*******************************************/

        private CommandBarControls AppendMenuItem(CommandBarControls menu, string name)
        {
            CommandBarPopup popup = menu.Add(MsoControlType.msoControlPopup, null, null, null, true) as CommandBarPopup;
            popup.Caption = name;
            return popup.Controls;
        }

        /*******************************************/
        /**** Private Fields                    ****/
        /*******************************************/

        private Dictionary<string, T> m_ItemLinks = new Dictionary<string, T>();
        private List<CommandBarControl> m_buttons = new List<CommandBarControl>();
    }
}
