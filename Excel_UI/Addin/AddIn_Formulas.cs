/*
 * This file is part of the Buildings and Habitats object Model (BHoM)
 * Copyright (c) 2015 - 2022, the respective contributors. All rights reserved.
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
using System.Linq.Expressions;
using Microsoft.Office.Interop.Excel;
using BH.oM.UI;
using BH.UI.Base;
using BH.UI.Excel.Templates;
using BH.oM.Base;
using BH.UI.Excel.Components;
using BH.oM.Versioning;

using BH.Engine.Serialiser;

namespace BH.UI.Excel
{
    public partial class AddIn : IExcelAddIn
    {
        /*******************************************/
        /**** Methods                           ****/
        /*******************************************/

        public static void Register(CallerFormula caller, System.Action callback = null, bool saveToHiddenSheet = true)
        {
            lock (m_Mutex)
            {
                // Register the caller formula with Excel if not already done
                if (m_Registered.Contains(caller.Function))
                {
                    if (callback != null)
                        ExcelAsyncUtil.QueueAsMacro(() => callback());
                }
                else
                {
                    var formula = caller.GetExcelDelegate();
                    string function = caller.Function;

                    ExcelAsyncUtil.QueueAsMacro(() =>
                    {
                        lock (m_Mutex)
                        {
                            if (!m_Registered.Contains(function))
                            {
                                ExcelIntegration.RegisterDelegates(
                                    new List<Delegate>() { formula.Item1 },
                                    new List<object> { formula.Item2 },
                                    new List<List<object>> { formula.Item3 }
                                );
                                m_Registered.Add(function);
                                ExcelDna.IntelliSense.IntelliSenseServer.Refresh();
                            }

                            if (callback != null)
                                ExcelAsyncUtil.QueueAsMacro(() => callback());
                        }
                    });
                }

                // Save the caller info to the hidden sheet if needed
                if (saveToHiddenSheet)
                    SaveCallerToHiddenSheet(caller);
            }
        }

        /*******************************************/

        public static void RestoreFormulas()
        {
            // Get the hidden worksheet
            Worksheet sheet = Sheet("BHoM_CallersHidden", false);
            if (sheet == null)
            {
                Old_Restore(); // is it an old version of an Excel file ?
                return;
            }

            // Get all the formulas stored in teh BHoM_CallersHidden sheet
            for (int i = 2; i < 10000; i++)
            {
                // Recover the information about the formula
                string formulaName = sheet.Cells[i, 1].Value as string;
                string callerJson = sheet.Cells[i, 2].Value as string;
                string oldFunction = sheet.Cells[i, 3].Value as string;
                if (formulaName == null || formulaName.Length == 0 || callerJson == null || callerJson.Length == 0)
                    break;

                // Register that formula from the json information
                CallerFormula formula = InstantiateCaller(formulaName);
                if (formula != null)
                {
                    BH.Engine.Base.Compute.ClearCurrentEvents();
                    formula.Caller.Read(callerJson);

                    VersioningEvent versioning = BH.Engine.Base.Query.CurrentEvents().OfType<VersioningEvent>().FirstOrDefault();

                    Register(formula, () =>
                    {
                        if (versioning != null && !string.IsNullOrEmpty(oldFunction))
                            UpgradeCellsFormula(formula, oldFunction);
                    });
                    
                }

                // Register the choices as objects if formula is a dropdown
                CallerValueListFormula valueList = formula as CallerValueListFormula;
                if (valueList != null)
                {
                    foreach (object choice in valueList.MultiChoiceCaller.Choices)
                    {
                        if (choice is IObject)
                            IAddObject(choice);
                    }

                    FUCKINGHELL(valueList.MultiChoiceCaller.SelectedItem.ToJson(), valueList.MultiChoiceCaller.Choices);
                }
            }
        }


        /*******************************************/
        /**** Private Methods                   ****/
        /*******************************************/

        private static string FUCKINGHELL(string collectionName, List<object> choices)
        {
            //Update enum values in case they have changed since the last serialisation
            var collectionNameObject = BH.Engine.Serialiser.Convert.FromJson(collectionName) as Type; //To strip out the 'BHoM_Version'
            var nameOfCollection = collectionNameObject.Namespace + "." + collectionNameObject.Name;

            var motherfuckingChoices = System.Enum.GetValues(collectionNameObject);

            // Get the data sheet
            Worksheet sheet = AddIn.Sheet("BHoM_ChoicesHidden", true, true);
            if (sheet == null)
                return string.Join(",", choices);

            // Try to find the list of choices in the spreadsheet
            int i = 0;
            while (i++ < 1000) // Just for safety
            {
                try
                {
                    string name = sheet.Cells[i, 1].Value as string;
                    if (string.IsNullOrEmpty(name))
                    {
                        // Need to add the choices here
                        sheet.Cells[i, 1].Value = collectionName;
                        for (int j = 0; j < choices.Count; j++)
                            sheet.Cells[i, j + 2].Value = choices[j];
                        break;
                    }
                    else
                    {
                        var sheetName = BH.Engine.Serialiser.Convert.FromJson(name) as Type;
                        var sheetNameObject = sheetName.Namespace + "." + sheetName.Name;
                        if (sheetNameObject == nameOfCollection)
                            break;
                    }
                }
                catch
                {
                    break;
                }
            }

            // Create the range
            Range range = sheet.Range[sheet.Cells[i, 2], sheet.Cells[i, choices.Count + 1]];

            for(int x = 0; x < motherfuckingChoices.Length; x++)
            {
                range[1, (x + 1)] = motherfuckingChoices.GetValue(x).ToString();
            }

            var rangeValidation = $"=BHoM_ChoicesHidden!{range.Address}";
            var rangeStart = $"=BHoM_ChoicesHidden!$B${i}:";
            
            Workbook workbook = ActiveWorkbook();
            if (workbook != null)
            {
                foreach (Worksheet sheety in workbook.Sheets.OfType<Worksheet>().Where(x => x.Visible == XlSheetVisibility.xlSheetVisible))
                {
                    var usedRange = sheety.UsedRange;

                    foreach(Range cell in usedRange)
                    {
                        Validation validation = cell.Validation;
                        if(validation != null)
                        {
                            try
                            {
                                string f1 = validation.Formula1;
                                string f2 = validation.Formula2;
                                if(f1.StartsWith(rangeStart) || f2.StartsWith(rangeStart))
                                {
                                    validation.Delete();
                                    validation.Add(XlDVType.xlValidateList, XlDVAlertStyle.xlValidAlertWarning, XlFormatConditionOperator.xlBetween, rangeValidation, Type.Missing);
                                    validation.InCellDropdown = true;
                                    validation.IgnoreBlank = true;
                                }
                            }
                            catch { }
                            /*validation.Delete();
                            validation.Add(XlDVType.xlValidateList, XlDVAlertStyle.xlValidAlertWarning, XlFormatConditionOperator.xlBetween, formula, Type.Missing);
                            validation.InCellDropdown = true;
                            validation.IgnoreBlank = true;*/
                        }
                    }
                }
            }

            return "fuck off";
            
            
        }

        private static void SaveCallerToHiddenSheet(CallerFormula caller)
        {
            // Get the hidden worksheet
            Worksheet sheet = Sheet("BHoM_CallersHidden", true, true);
            if (sheet == null)
                return;

            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                lock(m_Mutex)
                {
                    // Get teh sheet ID
                    string sheetId = sheet.Cells[1, 1].Value as string;
                    if (sheetId == null)
                    {
                        sheetId = ToString(Guid.NewGuid());
                        sheet.Cells[1, 1].Value = sheetId;
                    }

                    // Make sure there is a list of saved callers for this worksheet
                    if (!m_SavedOnWorkbook.ContainsKey(sheetId))
                        m_SavedOnWorkbook[sheetId] = new HashSet<string>();

                    // Save the caller if not already done
                    if (!m_SavedOnWorkbook[sheetId].Contains(caller.Function))
                    {
                        m_SavedOnWorkbook[sheetId].Add(caller.Function);
                        int row = m_SavedOnWorkbook[sheetId].Count + 1;

                        sheet.Cells[row, 1].Value = caller.Caller.GetType().Name;
                        sheet.Cells[row, 2].Value = caller.Caller.Write();
                        sheet.Cells[row, 3].Value = caller.Function;
                    } 
                } 
            });
            
        }

        /*******************************************/

        private static void UpgradeCellsFormula(CallerFormula formula, string oldFunction)
        {
            if (formula?.Caller?.SelectedItem == null)
                return;

            string oldFormula = '=' + oldFunction + '(';
            string newFormula = '=' + formula.Function + '(';

            Workbook workbook = ActiveWorkbook();
            if (workbook != null)
            {
                foreach (Worksheet sheet in workbook.Sheets.OfType<Worksheet>().Where(x => x.Visible == XlSheetVisibility.xlSheetVisible))
                    sheet.Cells.Replace(oldFormula, newFormula, XlLookAt.xlPart);
            }
        }


        /*******************************************/
        /**** Private Fields                    ****/
        /*******************************************/

        private static HashSet<string> m_Registered = new HashSet<string>();
        private static Dictionary<string, HashSet<string>> m_SavedOnWorkbook = new Dictionary<string, HashSet<string>>();
        private static object m_Mutex = new object();

        /*******************************************/
    }
}



