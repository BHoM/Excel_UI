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

namespace BH.UI.Excel
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [Guid("751262D0-CEF4-47BA-8A77-C9B349DE887A")]
    public class Server
    {
        /***************************************************/
        /**** Constructors                              ****/
        /***************************************************/

        public Server()
        {
            BH.Engine.Reflection.Compute.LoadAllAssemblies();
        }


        /***************************************************/
        /**** Public Methods                            ****/
        /***************************************************/

        public string SayHello()
        {
            return "Hello from the BHoM !";
        }

        /***************************************************/

        public Object CreateObject(string typeName)
        {
            Type type = BH.Engine.Reflection.Create.Type(typeName);
            if (type == null)
                return null;

            object instance = Activator.CreateInstance(type);
            return new Object(type, instance.PropertyDictionary());
        }

        /***************************************************/

        public Enum CreateEnum(string typeName, string value)
        {
            Type type = BH.Engine.Reflection.Create.Type(typeName);
            if (type == null)
                return null;

            object instance = Activator.CreateInstance(type);
            return new Enum(type, value);
        }

        /***************************************************/

        public object GetObject(string id)
        {
            object result = AddIn.GetObject(id);
            if (result == null)
                return null;
            else
                return result.IToCom();
        }

        /***************************************************/

        public string GetEnumName(string typeName, string value)
        {
            Type type = BH.Engine.Reflection.Create.Type(typeName);
            if (type == null)
                return "";

            return Engine.Excel.Compute.ParseEnum(type, value)?.ToString();
        }

        /***************************************************/

        public object CallMethod(string methodName, Collection inputs = null)
        {
            int index = methodName.LastIndexOf('.');
            if (index < 0)
                return null;

            string typeName = methodName.Substring(0, index);
            methodName = methodName.Substring(index + 1);

            Type type = BH.Engine.Reflection.Query.EngineTypeList().Where(x => x.FullName == typeName).FirstOrDefault();
            if (type == null)
                return null;

            List<MethodInfo> methods = type.GetMethods(BindingFlags.Public | BindingFlags.Static | BindingFlags.DeclaredOnly)
                    .Where(x => x.Name == methodName)
                    .OfType<MethodInfo>()
                    .ToList();

            return Helpers.RunBestComMethod(methods, inputs);
        }

        /***************************************************/

        public Adapter CreateAdapter(string adapterName, Collection inputs = null)
        {
            Type type = BH.Engine.Reflection.Query.AdapterTypeList().Where(x => x.FullName.Contains(adapterName)).FirstOrDefault();
            if (type == null)
                return null;

            return new Adapter(adapterName, inputs);
        }


        /***************************************************/

        public Adapter CreateAdapter(string adapterName, string filePath, Object toolkitConfig = null)
        {
            Type type = BH.Engine.Reflection.Query.AdapterTypeList().Where(x => x.FullName.Contains(adapterName)).FirstOrDefault();
            if (type == null)
                return null;

            return new Adapter(adapterName, filePath, toolkitConfig);
        }

        /***************************************************/
        /**** Private Methods                           ****/
        /***************************************************/



        /***************************************************/
    }

}
