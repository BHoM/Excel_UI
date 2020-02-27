/*
 * This file is part of the Buildings and Habitats object Model (BHoM)
 * Copyright (c) 2015 - 2020, the respective contributors. All rights reserved.
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
using BH.UI.Excel.Callers;
using BH.UI.Excel.Templates;
using BH.UI.Global;
using BH.UI.Templates;
using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace BH.UI.Excel.Global
{
    public class FormulaSearchMenu : SearchMenu
    {
        /*******************************************/
        /**** Constructors                      ****/
        /*******************************************/

        public FormulaSearchMenu(Dictionary<string, CallerFormula> callers) : base()
        {
            m_Callers = callers;
        }

        /*******************************************/
        /**** Methods                           ****/
        /*******************************************/

        public override bool SetParent(object parent)
        {
            List<Delegate> delegates = new List<Delegate>();
            List<ExcelFunctionAttribute> funcAttrs = new List<ExcelFunctionAttribute>();
            List<List<object>> argAttrs = new List<List<object>>();
            Dictionary<string, int> dups = new Dictionary<string, int>();
            foreach(var item in PossibleItems)
            {
                using (Engine.Excel.Profiling.Timer timer = new Engine.Excel.Profiling.Timer("CreateDelegates"))
                {
                    try
                    {
                        var proxy = CreateDelegate(item);
                        if (proxy == null)
                            continue;
                        var name = proxy.Item2.Name;
                        if (!dups.ContainsKey(name))
                        {
                            dups.Add(name, 1);
                            delegates.Add(proxy.Item1);
                            funcAttrs.Add(proxy.Item2);
                            argAttrs.Add(proxy.Item3);
                        }
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.Message);
                    }
                }
            }
            try
            {
                using (Engine.Excel.Profiling.Timer timer = new Engine.Excel.Profiling.Timer("RegisterDelegates"))
                {
                    ExcelIntegration.RegisterDelegates(delegates, funcAttrs.Cast<object>().ToList(), argAttrs);
                }
            } catch
            {
                return false;
            }
            return true;
        }

        /*******************************************/
        /**** Private Methods                   ****/
        /*******************************************/

        protected override List<SearchItem> GetAllPossibleItems()
        {
            var items = base.GetAllPossibleItems();

            // All Types
            items.AddRange(Engine.Reflection.Query.BHoMTypeList().Select(x => new SearchItem
            {
                Item = x, CallerType = typeof(CreateCustomCaller), Text = x.ToText(true)
            }));

            items.AddRange(Engine.UI.Query.CreateRequestItems().Select(x => new SearchItem {
                Item = x, CallerType = typeof(UI.Components.CreateRequestCaller), Text = x.ToText(true)
            }));

            return items;
        }
        
        /*******************************************/

        private Tuple<Delegate, ExcelFunctionAttribute, List<object>> CreateDelegate(SearchItem item)
        {
            if (m_Callers.ContainsKey(item.CallerType.Name))
            {
                CallerFormula caller = m_Callers[item.CallerType.Name];
                caller.Caller.SetItem(item.Item);
                FormulaDataAccessor accessor = caller.Caller.DataAccessor as FormulaDataAccessor;
                if(accessor != null)
                    return accessor.Wrap(caller, () => NotifySelection(item));
            }
            return null;
        }

        /*******************************************/
        /**** Private Fields                    ****/
        /*******************************************/

        private Dictionary<string, CallerFormula> m_Callers;

        /*******************************************/
    }
}

