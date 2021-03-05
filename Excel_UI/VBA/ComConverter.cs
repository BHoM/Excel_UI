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
using ExcelDna.ComInterop;
using BH.Engine.Reflection;
using System.Reflection;
using BH.oM.Base;
using System.Collections;

namespace BH.UI.Excel
{
    [ComVisible(false)]
    public static partial class ComConverter
    {
        /***************************************************/
        /**** ToCom Methods                             ****/
        /***************************************************/

        public static object ToCom(this object obj)
        {
            if (obj == null)
                return null;

            Type type = obj.GetType();

            if (obj is Enum || obj is Type || obj is MethodInfo)
                return obj.ToString();
            else if (obj is string || type.IsValueType)
                return obj;
            else if (obj is IObject)
                return ToCom(obj as IObject);
            else if (obj is IEnumerable)
                return ToCom(obj as IEnumerable);
            else
                return obj;
        }

        /***************************************************/

        public static Object ToCom(this IObject obj)
        {
            if (obj == null)
                return null;

            Type type = obj.GetType();
            Dictionary<string, object> properties = obj.PropertyDictionary().ToDictionary(x => x.Key, x => x.Value.ToCom());
            return new Object(type, properties);
        }

        /***************************************************/

        public static Collection ToCom(this IEnumerable obj)
        {
            if (obj == null)
                return null;

            return new Collection(obj.Cast<object>().Select(x => x.ToCom()));
        }

        /***************************************************/

        public static PushType ToCom(this BH.oM.Adapter.PushType obj)
        {
            PushType result = PushType.AdapterDefault;
            Enum.TryParse(obj.ToString(), out result);
            return result;
        }


        /***************************************************/
        /**** FromCom Methods                           ****/
        /***************************************************/

        public static object FromCom(this object obj)
        {
            if (obj == null)
                return null;

            Type type = obj.GetType();

            if (obj is Enum)
                return obj.ToString();
            else if (obj is string || type.IsValueType)
                return obj;
            else if (obj is Object)
                return FromCom(obj as Object);
            else if (obj is Collection)
                return FromCom(obj as Collection);
            else
                return obj;
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
                Engine.Reflection.Modify.SetPropertyValue(instance, propName, obj[propName].FromCom());
            //prop.AssignProperty(instance, obj[prop.Name].FromCom());

            return instance;
        }

        /***************************************************/

        public static List<object> FromCom(this Collection obj)
        {
            if (obj == null)
                return null;

            return obj.Cast<object>().Select(x => x.FromCom()).ToList();
        }

        /***************************************************/

        public static BH.oM.Adapter.PushType FromCom(this PushType obj)
        {
            BH.oM.Adapter.PushType result = oM.Adapter.PushType.AdapterDefault;
            Enum.TryParse(obj.ToString(), out result);
            return result;
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
