/*
 * This file is part of the Buildings and Habitats object Model (BHoM)
 * Copyright (c) 2015 - 2026, the respective contributors. All rights reserved.
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

using BH.oM.UI;
using BH.UI.Base;
using BH.UI.Base.Global;
using System;
using System.Security.Cryptography;
using System.Text;

namespace BH.UI.Excel.Global
{
    public static class CustomRibbon
    {
        /*******************************************/
        /**** Public Methods                    ****/
        /*******************************************/

        public static void Activate()
        {
            Initialisation.CustomRibbonEntryLoaded += OnCustomRibbonEntryLoaded;

            foreach (CustomRibbonEntry entry in Initialisation.CustomRibbonEntries)
                OnCustomRibbonEntryLoaded(null, entry);
        }

        /*******************************************/

        public static string DeriveId(CustomRibbonEntry entry)
        {
            string seed = $"{entry.TabName}|{entry.Category}|{entry.ItemJson}";
            using (MD5 md5 = MD5.Create())
                return "c" + BitConverter.ToString(md5.ComputeHash(Encoding.UTF8.GetBytes(seed)))
                                   .Replace("-", "").ToLowerInvariant();
        }


        /*******************************************/
        /**** Private Methods                   ****/
        /*******************************************/

        private static void OnCustomRibbonEntryLoaded(object sender, CustomRibbonEntry entry)
        {
            try
            {
                string id = DeriveId(entry);
                if (AddIn.CustomEntryShells.ContainsKey(id))
                    return;

                // Validate: confirm CallerType can be instantiated and ItemJson deserialises correctly.
                Caller temp = Activator.CreateInstance(entry.CallerType) as Caller;
                if (temp == null)
                {
                    BH.Engine.Base.Compute.RecordWarning($"Could not instantiate Caller for custom ribbon entry. Tab: {entry.TabName}, Category: {entry.Category}.");
                    return;
                }

                object item = BH.Engine.Serialiser.Convert.FromJson(entry.ItemJson);
                temp.SetItem(item);

                AddIn.CustomEntryShells[id] = entry;
            }
            catch (Exception e)
            {
                BH.Engine.Base.Compute.RecordWarning(e, $"Failed to register custom ribbon entry. Tab: {entry.TabName}, Category: {entry.Category}.");
            }
        }

        /*******************************************/
    }
}
