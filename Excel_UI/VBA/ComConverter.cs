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
using ExcelDna.ComInterop;
using BH.Engine.Reflection;
using System.Reflection;
using BH.oM.Base;
using System.Collections;
using BH.Engine.Base;

namespace BH.UI.Excel
{
    [ComVisible(false)]
    public static partial class ComConverter
    {
        /***************************************************/
        /**** ToCom Methods                             ****/
        /***************************************************/

        public static object IToCom(this object obj)
        {
            if (obj == null)
                return null;

            return ToCom(obj as dynamic);
        }

        /***************************************************/

        public static Object ToCom(this IObject obj)
        {
            if (obj == null)
                return null;

            Type type = obj.GetType();
            Dictionary<string, object> properties = obj.PropertyDictionary().ToDictionary(x => x.Key, x => x.Value.IToCom());

            if (obj is BHoMObject)
            {
                foreach (var kvp in ((BHoMObject)obj).CustomData)
                {
                    if (!properties.ContainsKey(kvp.Key))
                        properties.Add(kvp.Key, kvp.Value);
                }
            }

            return new Object(type, properties);
        }

        /***************************************************/

        public static string ToCom(this Enumeration enumeration)
        {
            return enumeration.Value;
        }

        /***************************************************/

        public static Collection ToCom(this FragmentSet obj)
        {
            return ToCom(obj as IEnumerable);
        }

        /***************************************************/

        public static Dictionary ToCom<T>(this Dictionary<string, T> obj)
        {
            if (obj == null)
                return null;

            return new Dictionary(obj.Keys, obj.Values.Cast<object>().Select(x => x.IToCom()));
        }

        /***************************************************/

        public static Collection ToCom(this IEnumerable obj)
        {
            if (obj == null)
                return null;

            return new Collection(obj.Cast<object>().Select(x => x.IToCom()));
        }

        /***************************************************/

        public static PushType ToCom(this BH.oM.Adapter.PushType obj)
        {
            PushType result = PushType.AdapterDefault;
            System.Enum.TryParse(obj.ToString(), out result);
            return result;
        }

        /***************************************************/

        public static string ToCom(this Type type)
        {
            return type?.ToText(true);
        }

        /***************************************************/

        public static string ToCom(this MethodBase method)
        {
            return method?.ToText(true);
        }

        /***************************************************/

        public static string ToCom(this string text)
        {
            return text;
        }

        /***************************************************/

        private static object ToCom(this object obj)
        {
            if (obj.GetType().IsEnum)
                return obj.ToString();
            else
                return obj;
        }


        /***************************************************/
        /**** FromCom Methods                           ****/
        /***************************************************/

        public static object IFromCom(this object obj)
        {
            if (obj == null || obj is DBNull)
                return null;

            return FromCom(obj as dynamic);
        }

        /***************************************************/

        public static object FromCom(this Object obj)
        {
            if (obj == null)
                return null;

            Type type = obj.GetCSharpType();
            if (type == null)
                return null;

            object instance = Activator.CreateInstance(type);
            foreach (string propName in obj.GetProperties())
                Engine.Base.Modify.SetPropertyValue(instance, propName, obj[propName].IFromCom());
            //prop.AssignProperty(instance, obj[prop.Name].FromCom());

            return instance;
        }

        /***************************************************/

        public static List<object> FromCom(this Collection obj)
        {
            if (obj == null)
                return null;

            return obj.Cast<object>().Select(x => x.IFromCom()).ToList();
        }

        /***************************************************/

        public static Dictionary<string, object> FromCom(this Dictionary obj)
        {
            if (obj == null)
                return null;

            return obj.Keys.Zip(obj.Values, (k, v) => new { k, v }).ToDictionary(x => x.k, x => x.v); 
        }

        /***************************************************/

        public static BH.oM.Adapter.PushType FromCom(this PushType obj)
        {
            BH.oM.Adapter.PushType result = oM.Adapter.PushType.AdapterDefault;
            System.Enum.TryParse(obj.ToString(), out result);
            return result;
        }

        /***************************************************/

        public static object FromCom(this Excel.Enum e)
        {
            if (e == null)
                return null;

            try
            {
                return BH.Engine.Excel.Compute.ParseEnum(e.GetCSharpType(), e.Value);
            }
            catch 
            {
                return null;
            }
        }

        /***************************************************/

        public static object FromCom(this string text)
        {
            return text;
        }

        /***************************************************/

        private static object FromCom(this object obj)
        {
            if (obj == null || obj is DBNull)
                return null;
            else
                return obj;
        }


        /***************************************************/
        /**** Private Methods                           ****/
        /***************************************************/

        private static void AssignProperty(this PropertyInfo prop, object instance, object value)
        {
            Type type = prop.PropertyType;
            if (value is IList && typeof(IList).IsAssignableFrom(type) && type.IsGenericType)
                value = TransferData(value as IList, Activator.CreateInstance(type) as dynamic);

            try
            {
                prop.SetValue(instance, value);
            }
            catch { }
        }

        /***************************************************/

        private static object TransferData<T>(IList source, IList<T> target)
        {
            return source.OfType<T>().ToList();
        }

        /***************************************************/
    }

}

