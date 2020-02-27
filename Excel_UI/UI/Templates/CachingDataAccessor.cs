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
            if(m_currentOp != null)
            {
                m_outputCache[m_currentOp] = base.GetOutput();
                m_cacheAge[m_currentOp] = DateTime.Now;
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

                m_currentOp = key;

                if (m_cacheAge.ContainsKey(m_currentOp))
                {
                    var age = DateTime.Now.Subtract(m_cacheAge[m_currentOp]);
                    if (age.TotalSeconds > 60)
                    {
                        m_cacheAge.Remove(m_currentOp);
                        m_outputCache.Remove(m_currentOp);
                        m_inputCache.Remove(m_currentOp);
                    }
                }

                if (function.Length > 0 && m_inputCache.ContainsKey(key))
                {
                    var cached = m_inputCache[key];
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
                m_inputCache[key] = in_;
            } catch { }

            return base.Store(function, in_);
        }

        /*******************************************/

        public override object GetOutput()
        {
            if (m_currentOp == null)
                return base.GetOutput();

            object output = null;
            m_outputCache.TryGetValue(m_currentOp, out output);

            base.SetDataItem(0, output);


            return base.GetOutput();
        }

        /*******************************************/
        /**** Private Fields                    ****/
        /*******************************************/

        private string m_currentOp;
        private Dictionary<string, object[]> m_inputCache = new Dictionary<string, object[]>();
        private Dictionary<string, object> m_outputCache = new Dictionary<string, object>();
        private Dictionary<string, DateTime> m_cacheAge = new Dictionary<string, DateTime>();

        /*******************************************/
    }
}

