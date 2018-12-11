using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BH.oM.Base;
using BH.oM.Geometry;
using BH.Engine.Reflection;
using BH.oM.DataManipulation.Queries;
using BH.Engine.Serialiser;
using Microsoft.Office.Interop.Excel;

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

        public bool Empty => m_objects.Count == 0;

        /*****************************************/
        /**** Public get methods        **********/
        /*****************************************/

        public IBHoMObject GetBHoM(string str)
        {
            return GetAny(str) as IBHoMObject;
        }

        /*****************************************/

        public IGeometry GetGeometry(string str)
        {
            return GetAny(str) as IGeometry;
        }

        /*****************************************/

        public object GetAny(string str)
        {
            if(m_objects.ContainsKey(str))
            {
                return m_objects[str]; 
            } else
            {
                int start = str.LastIndexOf("[");
                int end = str.LastIndexOf("]");
                if(start != -1 && end != -1 && end > start)
                {
                    return GetAny(str.Substring(++start, end - start));
                }
            }
            return null;
        }

        /*****************************************/


        public IQuery GetQuery(string str)
        {
            return GetAny(str) as IQuery;
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
            string guid = ToString(obj.BHoM_Guid);
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

        public string Add(IQuery query)
        {
            string guid = ToString(Guid.NewGuid());
            m_objects[guid] = query;
            return guid;
        }

        /*****************************************/

        public string Add(IGeometry geom)
        {
            string guid = ToString(Guid.NewGuid());
            m_objects[guid] = geom;
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

        public static Project ForWorkbook(Workbook Wb)
        {
            Project proj = new Project();
            foreach (Worksheet sheet in Wb.Sheets)
            {
                if (sheet.Name == "BHoM_Data") continue;
                foreach ( Range cell in sheet.UsedRange)
                {
                    try
                    {
                        if (cell.Value is string)
                        {
                            string val = cell.Value;
                            int start = val.LastIndexOf("[");
                            int end = val.LastIndexOf("]");
                            if (start != -1 && end != -1 && end > start)
                            {
                                ++start;
                                val = val.Substring(start, end - start);
                                object obj = Project.ActiveProject.GetAny(val);
                                if (obj != null)
                                {
                                    proj.m_objects.Add(val, obj);
                                }
                            }
                        }
                    }
                    catch { }
                }
            }
            return proj;
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
                    Engine.Reflection.Compute.RecordError(e.Message);
                }
            }
        }
    }
}
