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
using System.IO;
using System.Reflection;
using System.Linq;
using ExcelDna.Integration;
using System.Collections.Generic;
using System.Collections;
using NetOffice.ExcelApi;
using BH.UI.Excel.Templates;
using ExcelDna.Registration;

namespace BH.UI.Excel
{
    public partial class AddIn : IExcelAddIn
    {


        /*******************************************/
        /**** Methods                           ****/
        /*******************************************/

        public void AutoOpen()
        {
            // Install Excel DNA intelisense
            ExcelDna.IntelliSense.IntelliSenseServer.Install();

            // Initialise the BHoM
            ExcelAsyncUtil.QueueAsMacro(() => InitBHoMAddin());

            // Events on Excel itself
            Application app = Application.GetActiveInstance();
            if (app != null)
            {
                app.WorkbookOpenEvent += App_WorkbookOpen;
                app.WorkbookBeforeCloseEvent += App_WorkbookClosed;
            }
        }

        /*******************************************/

        public void AutoClose()
        {
            try
            {
                // note: This method only runs if the Addin gets disabled during
                // execution, it does not run when excel closes.
                ExcelDna.IntelliSense.IntelliSenseServer.Uninstall();

                Application app = Application.GetActiveInstance();
                if (app != null)
                {
                    app.WorkbookOpenEvent -= App_WorkbookOpen;
                    app.WorkbookBeforeCloseEvent -= App_WorkbookClosed;
                }
            }
            catch { }
        }


        /*******************************************/
        /**** Private Methods                   ****/
        /*******************************************/

        private void InitBHoMAddin()
        {
            // Make sure we initialise only once
            if (m_Initialised)
                return;
            m_Initialised = true;

            // Set up Excel DNA
            ExcelDna.Registration.ExcelRegistration.RegisterCommands(ExcelDna.Registration.ExcelRegistration.GetExcelCommands());
            ExcelDna.IntelliSense.IntelliSenseServer.Refresh();

            // Register single item formulas
            foreach (CallerFormula formula in CallerShells.Values.Where(x => !x.Caller.HasPossibleItems))
                Register(formula, null, false);

            // Register any function that was defined explicitely in this project
            ExcelRegistration.GetExcelFunctions().RegisterFunctions();

            // Initialise global search
            InitGlobalSearch();
            ExcelDna.Logging.LogDisplay.Clear();
        }

        /*******************************************/

        private void App_WorkbookOpen(Workbook workbook)
        {
            // Restore internalised data and callers
            RestoreData();
            RestoreFormulas();

            // Initialise the BHoM Addin and run first calculation
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                foreach (Worksheet sheet in workbook.Sheets.OfType<Worksheet>())
                {
                    bool before = sheet.EnableCalculation;
                    sheet.EnableCalculation = false;
                    sheet.Calculate();
                    sheet.EnableCalculation = before;
                }
            });
        }

        /*******************************************/

        private void App_WorkbookClosed(Workbook workbook, ref bool cancel)
        {
            ClearObjects();
        }


        /*******************************************/
        /**** Private Fields                    ****/
        /*******************************************/

        private bool m_Initialised = false;

        /*******************************************/
    }
}

