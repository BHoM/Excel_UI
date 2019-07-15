/*
 * This file is part of the Buildings and Habitats object Model (BHoM)
 * Copyright (c) 2015 - 2018, the respective contributors. All rights reserved.
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
using BH.oM.Base;
using BH.Engine.Reflection;
using BH.Engine.Serialiser;
using BH.Adapter;
using NetOffice.ExcelApi;
using NetOffice.ExcelApi.Enums;

namespace BH.UI.Excel
{
    //Singleton project class to hold the BHoM objects.
    //TODO: Should probably be moved to somewhere in the engine

    public class Project
    {
        /*****************************************/
        /**** Data fields               **********/
        /*****************************************/

        private Dictionary<string, object> m_objects;


        /*****************************************/
        /**** Static singleton instance **********/
        /*****************************************/

        private static Project m_instance = null;

        /*****************************************/
        /**** Constructor               **********/
        /*****************************************/

        private Project()
        {
            m_objects = new Dictionary<string, object>();
        }

        /*****************************************/
        /**** Get singleton method      **********/
        /*****************************************/

        public static Project ActiveProject
        {
            get
            {
                if (m_instance == null)
                    m_instance = new Project();

                return m_instance;
            }
        }

        public bool Empty => Count() == 0;

        /*****************************************/
        /**** Public get methods        **********/
        /*****************************************/

        public IBHoMObject GetBHoM(string str)
        {
            return GetAny(str) as IBHoMObject;
        }

        /*****************************************/

        public string GetId(string str)
        {
            if(m_objects.ContainsKey(str))
            {
                return str; 
            } else
            {
                int start = str.LastIndexOf("[");
                int end = str.LastIndexOf("]");
                if(start != -1 && end != -1 && end > start)
                {
                    return str.Substring(++start, end - start);
                }
            }
            return null;
        }

        public object GetAny(string str)
        {
            string id = GetId(str);
            if (id != null)
            {
                return m_objects[id];
            }
            return null;
        }


        /*****************************************/
        /****** "Interface" Add method     *******/
        /*****************************************/
        public string IAdd(object obj)
        {
            return Add(obj as dynamic);
        }

        /*****************************************/
        /***** Add methods             ***********/
        /*****************************************/

        public string Add(IBHoMObject obj)
        {
            string guid = ToString(Guid.NewGuid());
            if (m_objects.ContainsKey(guid))
                return guid;

            m_objects.Add(guid, obj);

            //Recurively add the objects dependecies
            foreach (object o in obj.PropertyObjects())
            {
                if (o is IBHoMObject)
                {
                    Add(o as IBHoMObject);
                } 
            }

            //Add all objects in the custom data
            foreach (KeyValuePair<string, object> kvp in obj.CustomData)
            {
                if (kvp.Value is IBHoMObject)
                {
                    Add(kvp.Value as IBHoMObject);
                }
            }
            return guid;
        }

        /*****************************************/

        private static string ToString(Guid id)
        {
            return System.Convert.ToBase64String(id.ToByteArray()).Remove(8);
        }

        /*****************************************/

        private string Add(object obj)
        {
            string guid = ToString(Guid.NewGuid());
            m_objects[guid] = obj;
            return guid;
        }

        /*****************************************/

        public static Project ForIDs(IEnumerable<string> ids)
        {
            Project proj = new Project();
            foreach (string id in ids)
            {
                try
                {
                    object obj = ActiveProject.GetAny(id);
                    if (obj != null)
                    {
                        proj.m_objects.Add(ActiveProject.GetId(id), obj);
                    }
                }
                catch { }
            }
            return proj;
        }

        /*****************************************/

        public int Count()
        {
            return m_objects.Count;
        }

        /*****************************************/

        public int Count(Func<object, bool> predicate)
        {
            return m_objects.Count((kvp) => predicate(kvp.Value));
        }

        /*****************************************/

        public IEnumerable<string> Serialize()
        {
            foreach(var kvp in m_objects)
            {
                string json = null;
                try
                {
                    if (kvp.Value is IBHoMObject)
                    {
                        json = kvp.Value.ToJson();
                    }
                    if (kvp.Value is BHoMAdapter)
                    {
                        // Don't serialise adapters, they don't deserialise
                        Compute.RecordWarning("BHoMAdapter types canned be serialised");
                        continue;
                    }
                    else
                    {
                        json = new Dictionary<string, object>()
                        {
                            { kvp.Key, kvp.Value }
                        }.ToJson();
                    }
                }
                catch { }
                if (json != null) yield return json;
            }
            yield break;
        }

        /*****************************************/

        public void Deserialize(IEnumerable<string> objs)
        {
            foreach (var str in objs)
            {
                try
                {
                    var obj = Engine.Serialiser.Convert.FromJson(str);
                    if (obj == null) continue;
                    if (obj is KeyValuePair<string, object>)
                    {
                        var kvp = (KeyValuePair<string, object>)obj;
                        m_objects.Add(kvp.Key, kvp.Value);
                    } else if (obj is CustomObject)
                    {
                        var co = obj as CustomObject;
                        foreach (var kvp in co.CustomData)
                        {
                            m_objects.Add(kvp.Key, kvp.Value);
                        }
                    } else if (obj is IBHoMObject)
                    {
                        Add(obj as IBHoMObject);
                    }
                }
                catch (Exception e)
                {
                    Compute.RecordError(e.Message);
                }
            }
        }

        /*****************************************/

        public void SaveData(Workbook Wb)
        {
            Worksheet newsheet;
            try
            {
                try
                {
                    newsheet = Wb.Sheets["BHoM_DataHidden"] as Worksheet;
                }
                catch
                {
                    // Backwards compatibility
                    newsheet = Wb.Sheets["BHoM_Data"] as Worksheet;
                }
            }
            catch
            {
                newsheet = Wb.Sheets.Add() as Worksheet;
            }
            newsheet.Name = "BHoM_DataHidden";
            newsheet.Visible = XlSheetVisibility.xlSheetHidden;

            int row = 0;
            foreach (string json in Serialize())
            {
                string contents = "";
                Range cell;
                do
                {
                    row++;
                    cell = newsheet.Cells[row, 1];
                    try
                    {
                        contents = cell.Value as string;
                    }
                    catch { }
                } while (contents != null && contents.Length > 0);

                int c = 0;
                while (c < json.Length)
                {
                    cell.Value = json.Substring(c);
                    c += (cell.Value as string).Length;
                    cell = cell.Next;
                }
            }
        }
    }
}
