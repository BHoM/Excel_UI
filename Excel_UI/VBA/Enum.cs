/*
 * This file is part of the Buildings and Habitats object Model (BHoM)
 * Copyright (c) 2015 - 2022, the respective contributors. All rights reserved.
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
using BH.Engine.Reflection;
using BH.oM.Base;
using System.Reflection;

namespace BH.UI.Excel
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [Guid("F01B959C-B6CD-4CA7-AA33-7E9A5191CE9A")]
    public class Enum
    {
        /***************************************************/
        /**** Properties                                ****/
        /***************************************************/

        public string Type
        {
            get
            {
                return GetTypeName();
            }

            set
            {
                SetType(value);
            }
        }

        /***************************************************/

        public string Value { get; set; } = "";


        /***************************************************/
        /**** Constructors                              ****/
        /***************************************************/

        public Enum()
        {

        }

        /***************************************************/

        public Enum(Type type, string value)
        {
            try
            {
                m_Type = type;
                Value = value;
            }
            catch { }
        }


        /***************************************************/
        /**** Public Methods                            ****/
        /***************************************************/

        public bool SetType(string typeName)
        {
            m_Type = BH.Engine.Base.Create.Type(typeName);

            if (m_Type != null)
            {
                System.Enum e = Activator.CreateInstance(m_Type) as System.Enum;
                Value = e.ToString();
            }

            return m_Type != null;
        }

        /***************************************************/

        public string GetTypeName()
        {
            if (m_Type != null)
                return m_Type.FullName;
            else
                return "";
        }

        /***************************************************/

        public Type GetCSharpType()
        {
            return m_Type;
        }

        /***************************************************/

        public void SetValue(string value)
        {
            Value = value;
        }

        /***************************************************/

        public string[] GetPossibleValues()
        {
            if (m_Type == null)
                return new string[0];
            else
                return System.Enum.GetNames(m_Type).ToArray();
        }


        /***************************************************/
        /**** Private Fields                            ****/
        /***************************************************/

        protected Type m_Type = null;

        /***************************************************/
    }

}

