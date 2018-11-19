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

        private Dictionary<Guid, object> m_objects;


        /*****************************************/
        /**** Static singleton instance **********/
        /*****************************************/

        private static Project m_instance = null;

        /*****************************************/
        /**** Constructor               **********/
        /*****************************************/

        private Project()
        {
            m_objects = new Dictionary<Guid, object>();
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

        public IBHoMObject GetBHoM(Guid guid)
        {
            object obj;
            if (m_objects.TryGetValue(guid, out obj))
                return (IBHoMObject)obj;
            else
                return null;
        }

        /*****************************************/

        public IBHoMObject GetBHoM(string str)
        {
            Guid guid;
            return Guid.TryParse(str, out guid) ? GetBHoM(guid) : null;
        }

        /*****************************************/

        public IGeometry GetGeometry(Guid guid)
        {
            object obj;
            if (m_objects.TryGetValue(guid, out obj))
                return (IGeometry)obj;
            else
                return null;
        }

        /*****************************************/

        public IGeometry GetGeometry(string str)
        {
            Guid guid;
            return Guid.TryParse(str, out guid) ? GetGeometry(guid) : null;
        }

        /*****************************************/

        public object GetAny(Guid guid)
        {
            object obj;
            if (m_objects.TryGetValue(guid, out obj))
                return obj;
            else
                return null;
        }

        /*****************************************/

        public object GetAny(string str)
        {
            Guid guid;
            return Guid.TryParse(str, out guid) ? GetAny(guid) : null;
        }

        /*****************************************/

        public BHoMAdapter GetAdapter(Guid guid)
        {
            object obj;
            if (m_objects.TryGetValue(guid, out obj))
                return (BHoMAdapter)obj;
            else
                return null;
        }

        /*****************************************/

        public BHoMAdapter GetAdapter(string str)
        {
            Guid guid;
            return Guid.TryParse(str, out guid) ? GetAdapter(guid) : null;
        }

        /*****************************************/

        public IQuery GetQuery(Guid guid)
        {
            object obj;
            if (m_objects.TryGetValue(guid, out obj))
                return (IQuery)obj;
            else
                return null;
        }

        /*****************************************/

        public IQuery GetQuery(string str)
        {
            Guid guid;
            return Guid.TryParse(str, out guid) ? GetQuery(guid) : null;
        }

        /*****************************************/
        /****** "Interface" Add method     *******/
        /*****************************************/
        public Guid IAdd(object obj)
        {
            return Add(obj as dynamic);
        }

        /*****************************************/
        /***** Add methods             ***********/
        /*****************************************/

        public Guid Add(IBHoMObject obj)
        {
            if (m_objects.ContainsKey(obj.BHoM_Guid))
                return obj.BHoM_Guid;

            Guid guid = obj.BHoM_Guid;
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
            return obj.BHoM_Guid;
        }

        /*****************************************/

        public Guid Add(BHoMAdapter adapter)
        {
            if (m_objects.ContainsKey(adapter.BHoM_Guid))
                return adapter.BHoM_Guid;

            m_objects[adapter.BHoM_Guid] = adapter;
            return adapter.BHoM_Guid;
        }

        /*****************************************/

        public Guid Add(IExcelObject excelObj)
        {
            if (m_objects.ContainsKey(excelObj.BHoM_Guid))
                return excelObj.BHoM_Guid;

            m_objects[excelObj.BHoM_Guid] = excelObj;
            return excelObj.BHoM_Guid;
        }

        /*****************************************/

        public Guid Add(IQuery query)
        {
            Guid guid = Guid.NewGuid();
            m_objects[guid] = query;
            return guid;
        }

        /*****************************************/

        public Guid Add(IGeometry geom)
        {
            Guid guid = Guid.NewGuid();
            m_objects[guid] = geom;
            return guid;
        }

        /*****************************************/

        private Guid Add(object obj)
        {
            Guid guid = Guid.NewGuid();
            m_objects[guid] = obj;
            return guid;
        }

        /*****************************************/
    }
}
