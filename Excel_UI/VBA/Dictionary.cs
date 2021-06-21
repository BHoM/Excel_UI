/*
 * This file is part of the Buildings and Habitats object Model (BHoM)
 * Copyright (c) 2015 - 2021, the respective contributors. All rights reserved.
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
 
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using System.Collections;

namespace BH.UI.Excel
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [Guid("E7FC13EE-7ACC-4898-8956-1A23A57A20CF")]
    public class Dictionary : IEnumerable
    {
        /***************************************************/
        /**** Properties                                ****/
        /***************************************************/

        public object this[string key]
        {
            get
            {
                if (m_Objects.ContainsKey(key))
                    return m_Objects[key];
                else
                    return null;
            }

            set
            {
                m_Objects[key] = value;
            }
        }

        /***************************************************/

        public string[] Keys
        {
            get
            {
                return m_Objects.Keys.ToArray();
            }
        }

        /***************************************************/

        public object[] Values
        {
            get
            {
                return m_Objects.Values.ToArray();
            }
        }


        /***************************************************/
        /**** Constructors                              ****/
        /***************************************************/

        public Dictionary()
        {

        }

        /***************************************************/

        public Dictionary(IEnumerable<string> keys, IEnumerable<object> values)
        {
            m_Objects = keys.Zip(values, (k, v) => new { k, v }).ToDictionary(x => x.k, x => x.v);
        }


        /***************************************************/
        /**** Public Methods                            ****/
        /***************************************************/

        public IEnumerator GetEnumerator()
        {
            return m_Objects.GetEnumerator();
        }

        /***************************************************/

        public void Add(string key, object value)
        {
            m_Objects[key] = value;
        }

        /***************************************************/

        public bool Remove(string key)
        {
            return m_Objects.Remove(key);
        }

        /***************************************************/

        public bool ContainsKey(string key)
        {
            return m_Objects.ContainsKey(key);
        }


        /***************************************************/
        /**** Private Fields                            ****/
        /***************************************************/

        protected Dictionary<string, object> m_Objects = new Dictionary<string, object>();

        /***************************************************/
    }

}
