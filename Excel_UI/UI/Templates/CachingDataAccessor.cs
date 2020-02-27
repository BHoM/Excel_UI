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

using BH.oM.Base;
using BH.oM.UI;
using BH.UI.Templates;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using System.Linq.Expressions;
using System.Reflection;
using BH.Engine.Reflection;
using BH.Engine.Excel;

namespace BH.UI.Excel.Templates
{
    public class CacheingDataAccessor : FormulaDataAccessor
    {
        /*******************************************/
        /**** Constructors                      ****/
        /*******************************************/

        public CacheingDataAccessor()
        {
        }

        /*******************************************/
        /**** Methods                           ****/
        /*******************************************/

        public override bool SetDataItem<T>(int index, T data)
        {
            bool success = base.SetDataItem(index, data);
            if(m_CurrentOp != null)
            {
                m_OutputCache[m_CurrentOp] = base.GetOutput();
                m_CacheAge[m_CurrentOp] = DateTime.Now;
            }
            return success;
        }
            
        /*******************************************/

        public override bool Store(string function, params object[] in_)
        {
            if (function == null)
                function = "";
            try
            {
                string reference = Engine.Excel.Query.Caller().RefText();
                string key = $"{reference}:::{function}";

                m_CurrentOp = key;

                if (m_CacheAge.ContainsKey(m_CurrentOp))
                {
                    var age = DateTime.Now.Subtract(m_CacheAge[m_CurrentOp]);
                    if (age.TotalSeconds > 60)
                    {
                        m_CacheAge.Remove(m_CurrentOp);
                        m_OutputCache.Remove(m_CurrentOp);
                        m_InputCache.Remove(m_CurrentOp);
                    }
                }

                if (function.Length > 0 && m_InputCache.ContainsKey(key))
                {
                    var cached = m_InputCache[key];
                    if (in_.Length == cached.Length)
                    {
                        bool same = true;
                        for (int i = 0; i < in_.Length; i++)
                        {
                            if (!in_[i].Equals(cached[i]))
                            {
                                same = false;
                                break;
                            }
                        }
                        if (same)
                            return false;
                    }
                }
                m_InputCache[key] = in_;
            } catch { }

            return base.Store(function, in_);
        }

        /*******************************************/

        public override object GetOutput()
        {
            if (m_CurrentOp == null)
                return base.GetOutput();

            object output = null;
            m_OutputCache.TryGetValue(m_CurrentOp, out output);

            base.SetDataItem(0, output);


            return base.GetOutput();
        }

        /*******************************************/
        /**** Private Fields                    ****/
        /*******************************************/

        private string m_CurrentOp;
        private Dictionary<string, object[]> m_InputCache = new Dictionary<string, object[]>();
        private Dictionary<string, object> m_OutputCache = new Dictionary<string, object>();
        private Dictionary<string, DateTime> m_CacheAge = new Dictionary<string, DateTime>();

        /*******************************************/
    }
}

