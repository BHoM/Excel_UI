using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BH.oM.Base;
using BH.oM.Geometry;
using BH.Engine.Reflection;
using BH.Adapter;
using BH.oM.DataManipulation.Queries;

namespace BH.UI.Dragon
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
            }
            return null;
        }

        /*****************************************/

        public BHoMAdapter GetAdapter(string str)
        {
            return GetAny(str) as BHoMAdapter;
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

        public string Add(BHoMAdapter adapter)
        {
            string guid = ToString(adapter.BHoM_Guid);
            if (m_objects.ContainsKey(guid))
                return guid;

            m_objects[guid] = adapter;
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
            return Convert.ToBase64String(id.ToByteArray()).Remove(8);
        }

        /*****************************************/

        private string Add(object obj)
        {
            string guid = ToString(Guid.NewGuid());
            m_objects[guid] = obj;
            return guid;
        }

        /*****************************************/
    }
}
