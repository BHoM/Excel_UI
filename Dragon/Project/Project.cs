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
        /**** Public methods            **********/
        /*****************************************/

        public Guid AddBHoM(IObject obj)
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
                    AddBHoM(o as BHoMObject);
                }
            }
            //Add all objects in the custom data
            foreach (KeyValuePair<string, object> kvp in obj.CustomData)
            {
                if (kvp.Value is BHoMObject)
                {
                    AddBHoM(kvp.Value as BHoMObject);
                }
            }
            return obj.BHoM_Guid;
        }

        /*****************************************/

        public IObject GetBHoM(Guid guid)
        {
            object obj;
            if (m_objects.TryGetValue(guid, out obj))
                return (IObject)obj;
            else
                return null;
        }

        /*****************************************/

        public IObject GetBHoM(string str)
        {
            Guid guid;
            return Guid.TryParse(str, out guid) ? GetBHoM(guid) : null;
        }

        /*****************************************/

        public Guid AddGeometry(IBHoMGeometry geom)
        {
            Guid guid = Guid.NewGuid();
            m_objects[guid] = geom;
            return guid;
        }

        /*****************************************/

        public IBHoMGeometry GetGeometry(Guid guid)
        {
            object obj;
            if (m_objects.TryGetValue(guid, out obj))
                return (IBHoMGeometry)obj;
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

        public Guid AddAdapter(BHoMAdapter adapter)
        {
            if (m_objects.ContainsKey(adapter.BHoM_Guid))
                return adapter.BHoM_Guid;

            m_objects[adapter.BHoM_Guid] = adapter;
            return adapter.BHoM_Guid;
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


        public Guid AddQuery(IQuery query)
        {
            Guid guid = Guid.NewGuid();
            m_objects[guid] = query;
            return guid;
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
    }
}
