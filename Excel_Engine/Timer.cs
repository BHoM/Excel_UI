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

using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BH.Engine.Excel.Profiling
{
    public class Timer : IDisposable
    {
        public Timer(string name)
        {
            m_name = name;
            m_start = DateTime.Now;
        }

        private string m_name;
        private DateTime m_start;
        
        private static Dictionary<string, List<double>> records = new Dictionary<string, List<double>>();

        private static void RecordTime(string name, double time)
        {
            if (records.ContainsKey(name))
            {
                records[name].Add(time);
            } else
            {
                records.Add(name, new List<double> { time });
            }
        }

        public static double GetTotal(string name)
        {
            if(records.ContainsKey(name))
            {
                return records[name].Sum();
            }
            return 0;
        }

        public static double GetMean(string name)
        {
            if(records.ContainsKey(name) && records[name].Count > 0)
            {
                return records[name].Sum() / records[name].Count;
            }
            return 0;
        }

        public void Dispose()
        {
            RecordTime(m_name, (DateTime.Now - m_start).TotalMilliseconds);
        }
    }
}
