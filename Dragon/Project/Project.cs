using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BH.oM.Base;
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

        private Dictionary<Guid, BHoMObject> m_objects;

        /*****************************************/
        /**** Static singleton instance **********/
        /*****************************************/

        private static Project m_instance = null;

        /*****************************************/
        /**** Constructor               **********/
        /*****************************************/

        private Project()
        {
            m_objects = new Dictionary<Guid, BHoMObject>();
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

        public void AddObject(BHoMObject obj)
        {
            if (m_objects.ContainsKey(obj.BHoM_Guid))
                return;

            Guid guid = obj.BHoM_Guid;
            m_objects.Add(guid, obj);

            foreach (object o in obj.PropertyObjects())
            {
                if (o is BHoMObject)
                {
                    AddObject(o as BHoMObject);
                }
            }
        }

        /*****************************************/

        public BHoMObject GetObject(Guid guid)
        {
            BHoMObject obj;
            if (m_objects.TryGetValue(guid, out obj))
                return obj;
            else
                return null;
        }

        /*****************************************/

        public BHoMObject GetObject(string str)
        {
            Guid guid;
            return Guid.TryParse(str, out guid) ? GetObject(guid) : null;
        }

    }
}
