using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BH.oM.Base;
using BH.oM.Geometry;
using BH.Engine.Reflection;
using BH.Adapter;
using BH.oM.Queries;

namespace BH.UI.Dragon
{
    //Singleton project class to hold the BHoM objects.
    //TODO: Should probably be moved to somewhere in the engine

    public class Project
    {
        /*****************************************/
        /**** Data fields               **********/
        /*****************************************/

        private Dictionary<Guid, IObject> m_objects;
        private Dictionary<Guid, IBHoMGeometry> m_geometry;
        private Dictionary<Guid, BHoMAdapter> m_adapters;
        private Dictionary<Guid, IQuery> m_queries;

        /*****************************************/
        /**** Static singleton instance **********/
        /*****************************************/

        private static Project m_instance = null;

        /*****************************************/
        /**** Constructor               **********/
        /*****************************************/

        private Project()
        {
            m_objects = new Dictionary<Guid, IObject>();
            m_geometry = new Dictionary<Guid, IBHoMGeometry>();
            m_adapters = new Dictionary<Guid, BHoMAdapter>();
            m_queries = new Dictionary<Guid, IQuery>();
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
        /**** Public methods            **********/
        /*****************************************/

        public Guid AddObject(IObject obj)
        {
            if (m_objects.ContainsKey(obj.BHoM_Guid))
                return obj.BHoM_Guid;

            Guid guid = obj.BHoM_Guid;
            m_objects.Add(guid, obj);

            //Recurively add the objects dependecies
            foreach (object o in obj.PropertyObjects())
            {
                if (o is BHoMObject)
                {
                    AddObject(o as BHoMObject);
                }
            }
            //Add all objects in the custom data
            foreach (KeyValuePair<string, object> kvp in obj.CustomData)
            {
                if (kvp.Value is BHoMObject)
                {
                    AddObject(kvp.Value as BHoMObject);
                }
            }
            return obj.BHoM_Guid;
        }

        /*****************************************/

        public IObject GetObject(Guid guid)
        {
            IObject obj;
            if (m_objects.TryGetValue(guid, out obj))
                return obj;
            else
                return null;
        }

        /*****************************************/

        public IObject GetObject(string str)
        {
            Guid guid;
            return Guid.TryParse(str, out guid) ? GetObject(guid) : null;
        }

        /*****************************************/

        public Guid AddGeometry(IBHoMGeometry geom)
        {
            Guid guid = Guid.NewGuid();
            m_geometry[guid] = geom;
            return guid;
        }

        /*****************************************/

        public IBHoMGeometry GetGeometry(Guid guid)
        {
            IBHoMGeometry obj;
            if (m_geometry.TryGetValue(guid, out obj))
                return obj;
            else
                return null;
        }

        /*****************************************/

        public IBHoMGeometry GetGeometry(string str)
        {
            Guid guid;
            return Guid.TryParse(str, out guid) ? GetGeometry(guid) : null;
        }

        /*****************************************/

        public object GetAny(Guid guid)
        {
            IObject obj;
            IBHoMGeometry geom;
            if (m_objects.TryGetValue(guid, out obj))
                return obj;
            else if (m_geometry.TryGetValue(guid, out geom))
                return geom;
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

        public Guid AddAdapter(BHoMAdapter adapter)
        {
            if (m_adapters.ContainsKey(adapter.BHoM_Guid))
                return adapter.BHoM_Guid;

            m_adapters[adapter.BHoM_Guid] = adapter;
            return adapter.BHoM_Guid;
        }

        /*****************************************/

        public BHoMAdapter GetAdapter(Guid guid)
        {
            BHoMAdapter obj;
            if (m_adapters.TryGetValue(guid, out obj))
                return obj;
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


        public Guid AddQuery(IQuery query)
        {
            Guid guid = Guid.NewGuid();
            m_queries[guid] = query;
            return guid;
        }

        /*****************************************/

        public IQuery GetQuery(Guid guid)
        {
            IQuery obj;
            if (m_queries.TryGetValue(guid, out obj))
                return obj;
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
    }
}
