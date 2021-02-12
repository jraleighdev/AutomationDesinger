using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using AutomationDesinger.Build.ApplicationFunctions;
using Microsoft.Office.Interop.Excel;
using SolidworksWrapper;
using SolidworksWrapper.Documents;
using AutomationDesinger.Constants;
using AutomationDesinger.Enums;
using AutomationDesinger.Helpers;
using System.Collections.Specialized;
using SolidworksWrapper.Constants;

namespace AutomationDesinger.Build
{
    public class ProcessRunBlockSolidworks : ExcelBaseParse
    {
        private SolidworksDocument _topDocument;

        private SolidworksDocument _workingDocument;

        private SolidworksMethods _methods;

        private bool inRepeat = false;
        private int repeatStart = 0;
        private int repeatEnd = 0;
        private int repeatCount = 0;
        private int repeatIndex = 0;

        public ProcessRunBlockSolidworks(Excel.Worksheet worksheet, SolidworksDocument topDocument = null) : base(worksheet)
        {
            if (!SolidworksApplication.Attached)
            {
                SolidworksApplication.Attach();
            }

            _methods = new SolidworksMethods();
        }

        public void Run(bool startMethod, string rangeName = "", string subRangeName = "", string subRangeValue = "")
        {
            Excel.Range startCell = null;

            if (string.IsNullOrEmpty(rangeName))
            {
                startCell = _worksheet.Range[_worksheet.Name + "Type"];
            }
            else
            {
                startCell = _worksheet.Range[rangeName];
            }

            if (!string.IsNullOrEmpty(subRangeName))
            {
                var subRangeCell = _worksheet.Range[subRangeName];
                subRangeCell.Value = subRangeValue;
            }

            if (startCell == null) throw new Exception("Could not find start range");


            var typeCol = startCell.Column;
            var nameCol = typeCol + 1;
            var parentCol = nameCol + 1;
            var value = parentCol + 1;
            var value2 = value + 1;
            var i = startCell.Row + 1;

            while (!string.IsNullOrEmpty(GetString(i, typeCol)))
            {
                var command = GetString(i, typeCol).ToUpper();

                if (command == Commands.Comment)
                {
                    i++;
                    continue;
                }

                var workingDocumentName = GetString(i, parentCol);

                if (command != Commands.Sub)
                {
                    if (string.IsNullOrEmpty(workingDocumentName))
                    {
                        if (_topDocument == null)
                        {
                            _topDocument = SolidworksApplication.ActiveDocument;
                        }
                        else
                        {
                            if (SolidworksApplication.ActiveDocument.Name != _topDocument.Name)
                            {
                                SolidworksApplication.ActiveDocument.Close();
                            }

                            SolidworksApplication.ActivateDocument(_topDocument.Name);
                        }

                        _workingDocument = _topDocument;
                    }
                    else
                    {
                        if (_workingDocument == null)
                        {
                            _workingDocument = SolidworksApplication.ActivateDocument(workingDocumentName);

                            if (_workingDocument == null)
                            {
                                throw new Exception($"Could not find document {workingDocumentName}");
                            }
                        }
                        else
                        {
                            if (_workingDocument.Name != workingDocumentName)
                            {
                                if (_workingDocument.Name != _topDocument.Name)
                                {
                                    _workingDocument.Close();
                                }

                                _workingDocument = SolidworksApplication.ActivateDocument(workingDocumentName);
                            }

                            if (_workingDocument == null)
                            {
                                throw new Exception($"Could not find document {workingDocumentName}");
                            }
                        }
                    }
                }

                switch (command)
                {
                    case Commands.TopLevelName:
                        var topLevelName = GetString(i, nameCol);

                        if (SolidworksApplication.ActiveDocument.Name != topLevelName)
                        {
                            throw new Exception("Top level name does not match active model");
                        }
                        else
                        {
                            _topDocument = SolidworksApplication.ActiveDocument;
                        }
                        break;
                    case Commands.Dimension:
                        _methods.SetDimValue(_workingDocument, GetString(i, nameCol), "", GetDouble(i, value));
                        break;
                    case Commands.Equation:
                        _methods.SetEquation(_workingDocument, GetString(i, nameCol), GetDouble(i, value));
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
                    case Commands.ShowConfiguration:
                        _methods.ShowConfiguration(_workingDocument, GetString(i, nameCol));
                        break;
                    case Commands.SetComponentConfiguration:
                        _methods.SetComponentConfiguration(_workingDocument, GetString(i, nameCol), GetString(i, value));
                        break;
                    case Commands.SetWeldmentConfiguration:
                        _methods.SetWeldmentMemberConfiguration(_workingDocument, GetString(i, nameCol), GetString(i, value));
                        break;
                    case Commands.ComponentVisiblity:
                        _methods.SetVisiblity(_workingDocument, GetString(i, nameCol), GetBoolean(i, value), FeatureTypes.Component);
                        break;
                    case Commands.DocumentReferenceVisiblity:
                        _methods.SetDocumentReferenceVisibility(_workingDocument, GetString(i, nameCol), GetBoolean(i, value), FeatureTypes.Component);
                        break;
                    case Commands.DeleteComponent:
                        // _methods.Delete(_workingDocument, GetString(i, nameCol));
                        break;
                    case Commands.DeleteReferencedDocuments:
                        // _methods.DeleteReferenced(_workingDocument, GetString(i, nameCol));
                        break;
                    case Commands.Stop:
                        throw new Exception("Program Stopped");
                    case Commands.UpdateDocument:
                        _workingDocument.ForceRebuildAll();
                        break;
                    case Commands.Sub:
                        ProcessRunBlockSolidworks runBlock = null;
                        var subName = GetString(i, nameCol);
                        var parameterName = GetString(i, parentCol);
                        var parameter = GetString(i, value);

                        var workSheetName = "";

                        if (subName.Contains("!"))
                        {
                            var subSplit = subName.Split('!');

                            workSheetName = subSplit[0];
                            subName = subSplit[1];

                            runBlock = new ProcessRunBlockSolidworks(Globals.ThisAddIn.Application.ActiveWorkbook.GetWorksheets().FirstOrDefault(x => x.Name == workSheetName), _topDocument);
                        }
                        else
                        {
                            runBlock = new ProcessRunBlockSolidworks(_worksheet, _topDocument);
                        }

                        runBlock.Run(false, subName, parameterName, parameter);
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

            if (startMethod)
            {
                if (_topDocument != null && SolidworksApplication.ActiveDocument.Name != _topDocument.Name)
                {
                    SolidworksApplication.ActiveDocument.Save();

                    SolidworksApplication.ActiveDocument.Close();
                }

                SolidworksApplication.ActiveDocument.ForceRebuildAll();

                SolidworksApplication.ActiveDocument.Save();
            }
        }
    }
}
