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
        /**** Public Methods                    ****/
        /*******************************************/

        public override bool SetDataItem<T>(int index, T data)
        {
            bool success = base.SetDataItem(index, data);
            if(current_op != null)
            {
                output_cache[current_op] = base.GetOutput();
                cache_age[current_op] = DateTime.Now;
            }
            return success;
        }

        public override bool Store(string function, params object[] in_)
        {
            if (function == null)
                function = "";
            try
            {
                string reference = Engine.Excel.Query.Caller().RefText();
                string key = $"{reference}:::{function}";

                current_op = key;

                if (cache_age.ContainsKey(current_op))
                {
                    var age = DateTime.Now.Subtract(cache_age[current_op]);
                    if (age.TotalSeconds > 60)
                    {
                        cache_age.Remove(current_op);
                        output_cache.Remove(current_op);
                        input_cache.Remove(current_op);
                    }
                }

                if (function.Length > 0 && input_cache.ContainsKey(key))
                {
                    var cached = input_cache[key];
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
                input_cache[key] = in_;
            } catch { }

            return base.Store(function, in_);
        }

        /*******************************************/

        public override object GetOutput()
        {
            if (current_op == null)
                return base.GetOutput();

            object output = null;
            output_cache.TryGetValue(current_op, out output);

            base.SetDataItem(0, output);


            return base.GetOutput();
        }

        /*******************************************/
        /**** Private Fields                    ****/
        /*******************************************/

        private string current_op;
        private Dictionary<string, object[]> input_cache = new Dictionary<string, object[]>();
        private Dictionary<string, object> output_cache = new Dictionary<string, object>();
        private Dictionary<string, DateTime> cache_age = new Dictionary<string, DateTime>();
    }
}

