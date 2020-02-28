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
        /*******************************************/
        /**** Properties                        ****/
        /*******************************************/

        public static Project ActiveProject
        {
            get
            {
                if (m_Instance == null)
                    m_Instance = new Project();

                return m_Instance;
            }
        }

        public bool Empty {
            get
            {
                return Count() == 0;
            }
        }

        /*******************************************/
        /**** Constructors                      ****/
        /*******************************************/

        private Project()
        {
            m_Objects = new Dictionary<string, object>();
        }

        /*******************************************/
        /**** Methods                           ****/
        /*******************************************/

        public IBHoMObject GetBHoM(string str)
        {
            return GetAny(str) as IBHoMObject;
        }

        /*******************************************/

        public string GetId(string str)
        {
            if(m_Objects.ContainsKey(str))
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

        /*******************************************/

        public object GetAny(string str)
        {
            string id = GetId(str);
            if (id != null)
            {
                return m_Objects[id];
            }
            return null;
        }

        /*******************************************/

        public string IAdd(object obj)
        {
            return IAdd(obj, ToString(Guid.NewGuid()));

        }

        /*******************************************/

        public string IAdd(object obj, Guid id)
        {
            return IAdd(obj, ToString(id));
        }

        /*****************************************/

        public string IAdd(object obj, string id)
        {
            return Add(obj as dynamic, id);
        }

        /*****************************************/

        public string Add(IBHoMObject obj, string id)
        {
            if (m_Objects.ContainsKey(id))
                return id;

            m_Objects.Add(id, obj);

            //Recurively add the objects dependecies
            foreach (object o in obj.PropertyObjects())
            {
                if (o is IBHoMObject)
                {
                    IAdd(o);
                }
            }

            //Add all objects in the custom data
            foreach (KeyValuePair<string, object> kvp in obj.CustomData)
            {
                if (kvp.Value is IBHoMObject)
                {
                    IAdd(kvp.Value);
                }
            }
            return id;
        }

        /*******************************************/

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
                        proj.m_Objects.Add(ActiveProject.GetId(id), obj);
                    }
                }
                catch { }
            }
            return proj;
        }

        /*******************************************/

        public int Count()
        {
            return m_Objects.Count;
        }

        /*******************************************/

        public int Count(Func<object, bool> predicate)
        {
            return m_Objects.Count((kvp) => predicate(kvp.Value));
        }

        /*******************************************/

        public IEnumerable<string> Serialize()
        {
            foreach(var kvp in m_Objects)
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
                        Engine.Reflection.Compute.RecordWarning("BHoMAdapter types canned be serialised");
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
                if (json != null)
                    yield return json;
            }
            yield break;
        }

        /*******************************************/

        public void Deserialize(IEnumerable<string> objs)
        {
            foreach (var str in objs)
            {
                try
                {
                    var obj = Engine.Serialiser.Convert.FromJson(str);
                    if (obj == null)
                        continue;
                    if (obj is KeyValuePair<string, object>)
                    {
                        var kvp = (KeyValuePair<string, object>)obj;
                        m_Objects.Add(kvp.Key, kvp.Value);
                    } else if (obj is IBHoMObject)
                    {
                        IAdd(obj, (obj as IBHoMObject).BHoM_Guid);
                    }
                }
                catch (Exception e)
                {
                    Engine.Reflection.Compute.RecordError(e.Message);
                }
            }
        }

        /*******************************************/

        public void SaveData(Workbook workbook)
        {
            Worksheet newsheet;
            try
            {
                try
                {
                    newsheet = workbook.Sheets["BHoM_DataHidden"] as Worksheet;
                }
                catch
                {
                    // Backwards compatibility
                    newsheet = workbook.Sheets["BHoM_Data"] as Worksheet;
                }
            }
            catch
            {
                newsheet = workbook.Sheets.Add() as Worksheet;
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

        /*******************************************/
        /**** Private Methods                   ****/
        /*******************************************/

        private string Add(object obj)
        {
            string guid = ToString(Guid.NewGuid());
            m_Objects[guid] = obj;
            return guid;
        }

        /*****************************************/

        private string Add(object obj, string id)
        {
            m_Objects[id] = obj;
            return id;
        }

        /*****************************************/

        private static string ToString(Guid id)
        {
            return System.Convert.ToBase64String(id.ToByteArray()).Remove(8);
        }


        /*****************************************/
        /**** Private Fields            **********/
        /*****************************************/

        private Dictionary<string, object> m_Objects;
        private static Project m_Instance = null;

        /*******************************************/
    }
}

