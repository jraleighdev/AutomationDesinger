using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using InventorWrapper.Documents;
using AutomationDesinger.Constants;
using InventorWrapper;
using AutomationDesinger.Build.ApplicationFunctions;
using System.Windows.Forms;
using System.Collections.Specialized;
using AutomationDesinger.Enums;
using AutomationDesinger.Helpers;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using InventorWrapper.Representation;
using AutomationDesinger.Logs;

namespace AutomationDesinger.Build
{
    public class ProcessRunBlockInventor : ExcelBaseParse
    {
        /// <summary>
        /// Top level document for application is running against
        /// </summary>
        private InventorDocument _topDocument;

        /// <summary>
        /// The current document we are working on
        /// </summary>
        private InventorDocument _workingDocument;

        /// <summary>
        /// Methods for manipulating inventor
        /// </summary>
        private InventorMethods _methods;

        /// <summary>
        /// If the application is currently in a repeat block
        /// </summary>
        private bool inRepeat = false;

        private int repeatStart = 0;
        private int repeatEnd = 0;
        private int repeatCount = 0;
        private int repeatIndex = 0;

        /// <summary>
        /// Attaches to inventor and setups up the methods
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="topDocument"></param>
        public ProcessRunBlockInventor(Excel.Worksheet worksheet, InventorDocument topDocument = null) : base(worksheet)
        {
            if (!InventorApplication.Attached)
            {
                InventorApplication.Attach();
            }

            _methods = new InventorMethods();
            _topDocument = topDocument;
        }

        /// <summary>
        /// Run application against the commands in the excel file
        /// </summary>
        /// <param name="rangeName"></param>
        /// <param name="rangeParameter"></param>
        /// <returns></returns>
        public void Run(string rangeName = "", string rangeParameter = "")
        {
            // The start cell
            Excel.Range startCell = null;

            // if the range name is null get the name of the current sheet plus type
            if (string.IsNullOrEmpty(rangeName))
            {
                startCell = _worksheet.Range[_worksheet.Name + "Type"];
            }
            else
            {
                startCell = _worksheet.Range[rangeName];
            }

            if (startCell == null) throw new Exception("Could not find start range");

            // Column for the type of commands we are running
            var typeCol = startCell.Column;

            // Name of the document or sub we are working with
            var nameCol = typeCol + 1;

            // Parent of the document we are working 
            var parentCol = nameCol + 1;

            // Value to set or get based on command
            var value = parentCol + 1;

            // Secondary value for some commands
            var value2 = value + 1;

            // Set the start row
            var i = startCell.Row + 1;

            // loop while the command column is not null
            while (!string.IsNullOrEmpty(GetString(i, typeCol)))
            {
                // get the command we are working with
                var command = GetString(i, typeCol).ToUpper();

                // if the command is a comment continue
                if (command == Commands.Comment)
                {
                    i++;
                    continue;
                }
                // Open the document 
                else if (command == Commands.OpenDocument)
                {
                    _methods.OpenDocument(GetString(i, nameCol), GetString(i, parentCol), GetString(i, value), GetString(i, value2));
                    i++;
                    continue;
                }

                // Assign the working document
                var workingDocumentName = GetString(i, parentCol);

                // working document name is null then assign the working document to the document
                if (string.IsNullOrEmpty(workingDocumentName))
                {
                    if (_topDocument == null)
                    {
                        _topDocument = InventorApplication.ActiveDocument;
                    }

                    _workingDocument = _topDocument;
                }
                // Assign the working document
                else
                {
                    // if the working document is null then assign
                    if (_workingDocument == null)
                    {
                        _workingDocument = DocumentHelper.GetDocument(workingDocumentName);

                        if (_workingDocument == null)
                        {
                            LogManager.Add($"Could not find document {workingDocumentName}. If all occurrences are suppressed you can ignore this error.");
                            i++;
                            continue;
                        }
                    }
                    // if the working document does not match working document name then assign it.
                    else
                    {
                        if (_workingDocument.Name != workingDocumentName)
                        {
                            _workingDocument = DocumentHelper.GetDocument(workingDocumentName);
                        }

                        if (_workingDocument == null)
                        {
                            LogManager.Add($"Could not find document {workingDocumentName}. If all occurrences are suppressed you can ignore this error.");
                            i++;
                            continue;
                        }
                    }
                }

                switch (command)
                {
                    case Commands.TopLevelName:
                        var topLevelName = GetString(i, nameCol);

                        if (InventorApplication.ActiveDocument.Name != topLevelName)
                        {
                            throw new Exception("Top level name does not match active model");
                        }
                        else
                        {
                            _topDocument = InventorApplication.ActiveDocument;
                        }
                        break;
                    case Commands.Parameter:
                        _methods.SetParameter(_workingDocument, GetString(i, nameCol), GetString(i, value));
                        break;
                    case Commands.GetParameter:
                        var parameterValue = _methods.GetParameter(_workingDocument, GetString(i, nameCol));
                        SetValue(i, value, parameterValue);
                        break;
                    case Commands.SetProperty:
                        _methods.SetProperty(_workingDocument, GetString(i, nameCol), GetString(i, value));
                        break;
                    case Commands.GetProperty:
                        var propertyValue = _methods.GetProperty(_workingDocument, GetString(i, nameCol));
                        SetValue(i, value, propertyValue);
                        break;
                    case Commands.ComponentActivity:
                        _methods.Suppression(_workingDocument, GetString(i, nameCol), GetString(i, value), SuppresionType.Component);
                        break;
                    case Commands.ConstraintActivity:
                        _methods.Suppression(_workingDocument, GetString(i, nameCol), GetString(i, value), SuppresionType.Constraint);
                        break;
                    case Commands.PatternActivity:
                        _methods.Suppression(_workingDocument, GetString(i, nameCol), GetString(i, value), SuppresionType.Pattern);
                        break;
                    case Commands.FeatureActivity:
                        _methods.Suppression(_workingDocument, GetString(i, nameCol), GetString(i, value), SuppresionType.Feature);
                        break;
                    case Commands.DeleteComponent:
                        _methods.Delete(_workingDocument, GetString(i, nameCol));
                        break;
                    case Commands.DeleteReferencedDocuments:
                        _methods.DeleteReferenced(_workingDocument, GetString(i, nameCol));
                        break;
                    case Commands.Stop:
                        throw new Exception("Program Stopped");
                    case Commands.UpdateDocument:
                        _workingDocument.Update();
                        break;
                    case Commands.Sub:
                        ProcessRunBlockInventor runBlock = null;
                        var subName = GetString(i, nameCol);
                        var workSheetName = "";
                        
                        if (subName.Contains("!"))
                        {
                            var subSplit = subName.Split('!');

                            workSheetName = subSplit[0];
                            subName = subSplit[1];

                            runBlock = new ProcessRunBlockInventor(Globals.ThisAddIn.Application.ActiveWorkbook.GetWorksheets().FirstOrDefault(x => x.Name == workSheetName), _topDocument);
                        }
                        else
                        {
                            runBlock = new ProcessRunBlockInventor(_worksheet, _topDocument);
                        }

                        runBlock.Run(subName, GetString(i, value));
                        break;
                    case Commands.If:
                        ValidateIf(i, typeCol);

                        var booleanValue = GetBoolean(i, value);

                        if (!booleanValue)
                        {
                            i = GetEndIfRow(i, typeCol);
                        }
                        break;
                    case Commands.Repeat:
                        ValidateRepeat(i, typeCol);
                        repeatStart = i;
                        repeatEnd = GetEndRepeatRow(i, typeCol);
                        repeatCount = GetInt(i, value);
                        repeatIndex = 1;
                        SetValue(i, value2, repeatIndex.ToString());
                        inRepeat = true;
                        break;
                }

                i++;

                if (inRepeat)
                {
                    if (i == repeatEnd)
                    {
                        if (repeatIndex == repeatCount)
                        {
                            i = repeatEnd + 1;
                            repeatStart = 0;
                            repeatEnd = 0;
                            repeatCount = 0;
                            repeatIndex = 0;
                            inRepeat = false;
                        }
                        else
                        {
                            i = repeatStart + 1;
                            repeatIndex++;
                            SetValue(repeatStart, value2, repeatIndex.ToString());
                        }
                    }
                }
            }

            if (_topDocument != null)
            {
                if (InventorApplication.ActiveDocument.Name != _topDocument.Name)
                {
                    InventorApplication.ActiveDocument.Save();

                    InventorApplication.ActiveDocument.Close();
                }
            }

            InventorApplication.ActiveDocument.Update();

            InventorApplication.ActiveDocument.Save();
        }
    }
}
