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
        /*******************************************/
        /**** Constructors                      ****/
        /*******************************************/

        public Timer(string name)
        {
            m_Name = name;
            m_Start = DateTime.Now;
        }

        /*******************************************/
        /**** Methods                           ****/
        /*******************************************/

        public static double GetTotal(string name)
        {
            if(m_Records.ContainsKey(name))
            {
                return m_Records[name].Sum();
            }
            return 0;
        }

        /*******************************************/

        public static double GetMean(string name)
        {
            if(m_Records.ContainsKey(name) && m_Records[name].Count > 0)
            {
                return m_Records[name].Sum() / m_Records[name].Count;
            }
            return 0;
        }

        /*******************************************/

        public void Dispose()
        {
            RecordTime(m_Name, (DateTime.Now - m_Start).TotalMilliseconds);
        }

        /*******************************************/
        /**** Private Methods                   ****/
        /*******************************************/

        private static void RecordTime(string name, double time)
        {
            if (m_Records.ContainsKey(name))
            {
                m_Records[name].Add(time);
            } else
            {
                m_Records.Add(name, new List<double> { time });
            }
        }

        /*******************************************/
        /**** Private Fields                    ****/
        /*******************************************/

        private string m_Name;
        private DateTime m_Start;
        private static Dictionary<string, List<double>> m_Records = new Dictionary<string, List<double>>();

        /*******************************************/
    }
}
