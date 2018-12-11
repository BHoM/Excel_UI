using System;
using System.Collections.Generic;
using System.Reflection;

namespace BH.Engine.Reflection
{
    //TODO: This whole class is copied from the Reflection_Engine and modified slightly to make the m_BHoMTypeList include interfacetypes as well as objecttypes.
    //This should not sit here but be part of the reflection engine that would need some additional tweaks.

    public static partial class ReflectionExtra
    {
        /***************************************************/
        /**** Public Methods                            ****/
        /***************************************************/

        public static List<Type> BHoMTypeList()
        {
            // If the dictionary exists already return it
            if (m_BHoMTypeList != null && m_BHoMTypeList.Count > 0)
                return m_BHoMTypeList;

            // Otherwise, create it
            ExtractAllTypes();

            return m_BHoMTypeList;
        }

        /***************************************************/

        public static List<Type> AdapterTypeList()
        {
            // If the dictionary exists already return it
            if (m_AdapterTypeList != null && m_AdapterTypeList.Count > 0)
                return m_AdapterTypeList;

            // Otherwise, create it
            ExtractAllTypes();

            return m_AdapterTypeList;
        }


        /***************************************************/
        /**** Private Methods                           ****/
        /***************************************************/

        private static void ExtractAllTypes()
        {
            m_BHoMTypeList = new List<Type>();
            m_AdapterTypeList = new List<Type>();

            foreach (Assembly asm in AppDomain.CurrentDomain.GetAssemblies())
            {
                try
                {
                    // Save BHoM objects only
                    string name = asm.GetName().Name;
                    if (name == "BHoM" || name.EndsWith("_oM"))
                    {
                        foreach (Type type in asm.GetTypes())
                        {
                            if (type.Namespace != null && type.Namespace.StartsWith("BH.oM"))
                            {
                                m_BHoMTypeList.Add(type);
                            }
                        }
                    }
                    // Save adapters
                    else if (name.EndsWith("_Adapter"))
                    {
                        foreach (Type type in asm.GetTypes())
                        {
                            if (!type.IsInterface)
                                m_AdapterTypeList.Add(type);
                        }
                    }
                }
                catch (Exception)
                {
                    Console.WriteLine("Cannot load types from assembly " + asm.GetName().Name);
                }
            }
        }


        /***************************************************/
        /**** Private Fields                            ****/
        /***************************************************/

        private static List<Type> m_BHoMTypeList = new List<Type>();
        private static List<Type> m_AdapterTypeList = new List<Type>();


        /***************************************************/
    }
}
