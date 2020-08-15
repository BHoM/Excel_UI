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

using BH.Adapter;
using BH.Engine.Reflection;
using BH.Engine.UI;
using BH.oM.UI;
using BH.UI.Base.Components;
using BH.UI.Excel.Templates;
using BH.UI.Base.Global;
using BH.UI.Base;
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
            var deduplicated = new Dictionary<string, SearchItem>();
            foreach(var item in PossibleItems)
            {
                try
                {
                    string fn = GetFormula(item);
                    if (deduplicated.ContainsKey(fn))
                    {
                        if (deduplicated[fn].Item.IIsDeprecated())
                        {
                            deduplicated[fn] = item;
                        }
                        continue;
                    }

                    deduplicated.Add(fn, item);
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }
            }

            foreach (var item in deduplicated.Values)
            {
                try
                {
                    RegisterDelegate(item);
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }
            }
            CallerFormula.RegisterQueue();
            ExcelDna.IntelliSense.IntelliSenseServer.Refresh();
            ExcelDna.Logging.LogDisplay.Hide();
            return true;
        }

        /*******************************************/
        /**** Private Methods                   ****/
        /*******************************************/

        private void RegisterDelegate(SearchItem item)
        {
            if (m_Callers.ContainsKey(item.CallerType.Name))
            {
                CallerFormula caller = m_Callers[item.CallerType.Name];
                caller.Caller.SetItem(item.Item);
                caller.EnqueueRegistration();
            }
        }
        
        /*******************************************/

        private string GetFormula(SearchItem item)
        {
            if (m_Callers.ContainsKey(item.CallerType.Name))
            {
                CallerFormula caller = m_Callers[item.CallerType.Name];
                caller.Caller.SetItem(item.Item);
                return caller.Function;
            }
            return null;
        }

        protected override List<SearchItem> GetAllPossibleItems()
        {
            var items = base.GetAllPossibleItems();

            // All constructors for the BHoM objects

            items.AddRange(Engine.Reflection.Query.BHoMTypeList()
                .Where(x => !x.IsNotImplemented() && x.IsDeprecated() && x?.GetInterface("IImmutable") == null && !x.IsEnum && !x.IsAbstract)
                .Select(x => new SearchItem { Item = x, CallerType = typeof(CreateObjectCaller), Text = x.ConstructorText() }));

            // All methods for the BHoM Engine
            items.AddRange(Engine.Reflection.Query.BHoMMethodList().Where(x=>x.IsDeprecated())
                .Select(x => new SearchItem { Item = x, CallerType = GetCallerType(x), Icon = GetIcon(x), Text = x.ToText(true) }));

            // All adapter constructors
            items.AddRange(Engine.Reflection.Query.AdapterTypeList()
                .Where(x => x.IsSubclassOf(typeof(BHoMAdapter)))
                .SelectMany(x => x.GetConstructors())
                .Where(x => !x.IsNotImplemented() && x.IsDeprecated())
                .Select(x => new SearchItem { Item = x, CallerType = typeof(CreateAdapterCaller), Text = x.ToText(true) }));

            // All Types
            items.AddRange( Engine.Reflection.Query.BHoMTypeList()
                .Concat(Engine.Reflection.Query.BHoMInterfaceList())
                .Where(x => !x.IsNotImplemented() && x.IsDeprecated())
                .Select(x => new SearchItem { Item = x, CallerType = typeof(CreateTypeCaller), Text = x.ToText(true) }));

            // All Enums
            items.AddRange(Engine.Reflection.Query.BHoMEnumList()
                .Where(x => !x.IsNotImplemented() && x.IsDeprecated())
                .Select(x => new SearchItem { Item = x, CallerType = typeof(CreateEnumCaller), Text = x.ToText(true) }));

            return items;
        }

        /*************************************/

        private static Type GetCallerType(MethodBase item)
        {
            if (item.DeclaringType.Namespace.StartsWith("BH.Engine"))
            {
                switch (item.DeclaringType.Name)
                {
                    case "Compute":
                        return typeof(ComputeCaller);
                    case "Convert":
                        return typeof(ConvertCaller);
                    case "Create":
                        return typeof(CreateObjectCaller);
                    case "Modify":
                        return typeof(ModifyCaller);
                    case "Query":
                        return typeof(QueryCaller);
                    default:
                        return null;
                }
            }
            else
                return null;
        }


        /*******************************************/
        /**** Private Fields                    ****/
        /*******************************************/

        private Dictionary<string, CallerFormula> m_Callers;

        /*******************************************/
    }
}

