using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BH.oM.Base;
using BH.oM.Geometry;
using BH.Engine.Reflection;

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

        public void AddObject(IObject obj)
        {
            if (m_objects.ContainsKey(obj.BHoM_Guid))
                return;

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
    }
}
