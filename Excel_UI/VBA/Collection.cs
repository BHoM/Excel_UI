/*
 * This file is part of the Buildings and Habitats object Model (BHoM)
 * Copyright (c) 2015 - 2023, the respective contributors. All rights reserved.
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
    [Guid("295011AF-BDEE-4E59-A6C3-AF4996B6414C")]
    public class Collection : IEnumerable
    {
        /***************************************************/
        /**** Properties                                ****/
        /***************************************************/

        public object this[int index]
        {
            get
            {
                if (index >= 0 && index < m_Objects.Count)
                    return m_Objects[index];
                else
                    return null;
            }

            set
            {
                if (index >= 0 && index < m_Objects.Count)
                    m_Objects[index] = value;
            }
        }

        /***************************************************/

        public int Count
        {
            get
            {
                return m_Objects.Count;
            }
        }


        /***************************************************/
        /**** Constructors                              ****/
        /***************************************************/

        public Collection()
        {

        }

        /***************************************************/

        public Collection(IEnumerable<object> items)
        {
            m_Objects = items.ToList();
        }


        /***************************************************/
        /**** Public Methods                            ****/
        /***************************************************/

        public IEnumerator GetEnumerator()
        {
            return m_Objects.GetEnumerator();
        }

        /***************************************************/

        public void Add(object item)
        {
            m_Objects.Add(item);
        }

        /***************************************************/

        public bool Remove(object item)
        {
            return m_Objects.Remove(item);
        }


        /***************************************************/
        /**** Private Fields                            ****/
        /***************************************************/

        protected List<object> m_Objects = new List<object>();

        /***************************************************/
    }

}


