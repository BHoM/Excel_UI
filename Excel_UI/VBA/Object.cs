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
using BH.Engine.Reflection;
using BH.oM.Base;
using System.Reflection;

namespace BH.UI.Excel
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [Guid("F44A27DD-7006-43AC-824E-82595BB75DB4")]
    public class Object
    {
        /***************************************************/
        /**** Properties                                ****/
        /***************************************************/

        public object this[string key]
        {
            get
            {
                return GetProperty(key);
            }

            set
            {
                SetProperty(key, value);
            }
        }

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

        public string[] PropertyNames
        {
            get
            {
                return m_Properties.Keys.ToArray();
            }
        }

        /***************************************************/

        public object[] PropertyValues
        {
            get
            {
                return m_Properties.Values.ToArray();
            }
        }


        /***************************************************/
        /**** Constructors                              ****/
        /***************************************************/

        public Object()
        {

        }

        /***************************************************/

        public Object(Type type, Dictionary<string, object> properties)
        {
            try
            {
                m_Type = type;
                m_Properties = properties;
            }
            catch { }
        }


        /***************************************************/
        /**** Public Methods                            ****/
        /***************************************************/

        public void SetProperty(string key, object value)
        {
            if (key == "_t" && value is string)
                SetType(value as string);
            else
                m_Properties[key] = value;
        }

        /***************************************************/

        public object GetProperty(string key)
        {
            if (m_Properties.ContainsKey(key))
                return m_Properties[key];
            else
                return null;
        }

        /***************************************************/

        public string[] GetProperties()
        {
            return m_Properties.Keys.ToArray();
        }

        /***************************************************/

        public bool SetType(string typeName)
        {
            m_Type = BH.Engine.Base.Create.Type(typeName);

            if (m_Type != null)
            {
                object instance = Activator.CreateInstance(m_Type);
                m_Properties = m_Type.GetProperties().ToDictionary(x => x.Name, x => x.GetValue(instance));
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

        public string GetPropertyType(string propertyName)
        {
            if (m_Type == null)
                return "";

            try
            {
                PropertyInfo prop = m_Type.GetProperty(propertyName);
                if (prop == null)
                    return "";
                else
                    return prop.PropertyType.FullName;
            }
            catch
            {
                return "";
            }
        }

        /***************************************************/

        public string GetBHoMId()
        {
            if (!m_Properties.ContainsKey("BHoM_Guid"))
                m_Properties["BHoM_Guid"] = Guid.NewGuid();

            return m_Properties["BHoM_Guid"].ToString();
        }

        /***************************************************/

        public void SetBHoMId(string id)
        {
            m_Properties["BHoM_Guid"] = Guid.Parse(id);
        }

        /***************************************************/

        public Type GetCSharpType()
        {
            return m_Type;
        }


        /***************************************************/
        /**** Private Fields                            ****/
        /***************************************************/

        protected Dictionary<string, object> m_Properties = new Dictionary<string, object>();

        protected Type m_Type = null;

        /***************************************************/
    }

}


