using InventorWrapper;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading.Tasks;
using AutomationDesinger.Helpers;
using AutomationDesinger.Constants;
using System.Linq;
using System.Collections.Generic;
using AutomationDesinger.DTOS;
using Microsoft.Office.Interop.Excel;
using System;
using AutomationDesinger.Build;
using Microsoft.Office.Core;
using System.Windows.Forms;
using AutomationDesinger.Forms;
using AutomationDesinger.CopyTools;
using InventorWrapper.CopyTools;
using AutomationDesinger.Enums;
using System.ComponentModel.Design;
using AutomationDesinger.Forms.SubForms;
using Syncfusion.Data.Extensions;
using System.IO;
using SolidworksWrapper;
using SolidworksWrapper.Enums;
using SolidworksWrapper.CopyTools;
using AutomationDesinger.Forms.DrawingCapture;
using AutomationDesinger.Logs;

namespace AutomationDesinger
{
    public partial class InventorBuildRibbon
    {
        private ProcessRunBlockInventor _processInventor;

        private ProcessRunBlockSolidworks _processSolidworks;

        private ApplicationTypeEnum ApplicationType;

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            InventorStopButton.Enabled = false;
            solidWorksStopBuild.Enabled = false;
        }

        #region Inventor Methods

        private void InventorBuildButton_Click(object sender, RibbonControlEventArgs e)
        {
            _processInventor = new ProcessRunBlockInventor(Globals.ThisAddIn.Application.ActiveSheet);

            try
            {
                _processInventor.Run();

                if (LogManager.HasData)
                {
                    LogManager.WriteLogs(Globals.ThisAddIn.Application.ActiveWorkbook);

                    MessageBox.Show("Done", "Application Complete - Please Review Logs", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
                else
                {
                    MessageBox.Show("Done", "Application Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ran into an error while running", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                LogManager.Clear();
                InventorStopButton.Enabled = false;
            }
        }

        private void InventorGetRefButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!InventorApplication.Attached)
            {
                InventorApplication.Attach();
            }

            var activeDocument = InventorApplication.ActiveDocument;

            var valid = activeDocument != null && activeDocument.IsAssemblyDoc;

            if (!valid)
            {
                MessageBox.Show("Please open an assembly document");
            }

            var adoc = activeDocument.GetAssemblyDocument();

            var docs = new List<CopyDocumentItem>();

            docs.Add(new CopyDocumentItem(adoc));

            var filesPathsToAvoid = Settings.Default.PathsToAvoid.SettingsToList();

            foreach (var doc in adoc.ReferencedDocuments)
            {
                if (filesPathsToAvoid.Any(x => doc.FileName.ToUpper().Contains(x.ToUpper()))) continue;

                docs.Add(new CopyDocumentItem(doc));
            }

            if (Globals.ThisAddIn.Application.ActiveWorkbook.WorkSheetExists("Copy Tool"))
            {
                Globals.ThisAddIn.Application.ActiveWorkbook.ActivateSheet("Copy Tool");
            }
            else
            {
                Globals.ThisAddIn.Application.ActiveWorkbook.Sheets.Add();

                Globals.ThisAddIn.Application.ActiveSheet.Name = "Copy Tool";
            }

            Excel.Worksheet workSheet = Globals.ThisAddIn.Application.ActiveSheet;

            workSheet.Cells.Clear();

            workSheet.Name = "Copy Tool";

            workSheet.Range["A1"].Value = "Old Path";
            workSheet.Range["B1"].Value = "New Path";

            var i = 2;

            foreach (var d in docs)
            {
                workSheet.Range[$"A{i}"].Value = d.OldPath;
                workSheet.Range[$"B{i}"].Value = d.NewPath;

                i++;
            }

            var copyTable = workSheet.Range[$"A1:B{i}"];

            workSheet.ListObjects.AddEx(XlListObjectSourceType.xlSrcRange, copyTable, null, XlYesNoGuess.xlYes).Name = "Copy Table";

            var outTable = workSheet.GetListObjects().FirstOrDefault(x => x.Name == "Copy Table");

            var old = outTable.GetListColumns().FirstOrDefault(x => x.Name == "Old Path");
        }

        private void InventorcopyDocuments_Click(object sender, RibbonControlEventArgs e)
        {
            if (!InventorApplication.Attached)
            {
                InventorApplication.Attach();
            }

            var activeDocument = InventorApplication.ActiveDocument;

            var valid = activeDocument != null && activeDocument.IsAssemblyDoc;

            if (!valid)
            {
                MessageBox.Show("Please open an assembly document");
            }

            var adoc = activeDocument.GetAssemblyDocument();

            if (Globals.ThisAddIn.Application.ActiveWorkbook.WorkSheetExists("Copy Tool"))
            {
                Globals.ThisAddIn.Application.ActiveWorkbook.ActivateSheet("Copy Tool");
            }
            else
            {
                return;
            }

            Excel.Worksheet workSheet = Globals.ThisAddIn.Application.ActiveSheet;

            try
            {
                var outTable = workSheet.GetListObjects().FirstOrDefault(x => x.Name == "Copy Table");

                if (outTable == null) throw new Exception("Could not find table Copy Table");

                var oldPath = outTable.GetListColumns().FirstOrDefault(x => x.Name == "Old Path");

                var newPath = outTable.GetListColumns().FirstOrDefault(x => x.Name == "New Path");

                var oldPathList = oldPath.Range.GetFilePaths();

                var newPathList = newPath.Range.GetFilePaths();

                if (oldPathList.Count != newPathList.Count)
                {
                    throw new Exception("Old path list and new path list do not match in qty");
                }

                var paths = new List<Tuple<string, string>>();

                for (var i = 0; i < oldPathList.Count; i++)
                {
                    paths.Add(new Tuple<string, string>(oldPathList[i], newPathList[i]));
                }

                var mainDocPath = paths.FirstOrDefault(x => x.Item1.ToUpper() == adoc.FileName.ToUpper());

                if (!System.IO.Directory.Exists(System.IO.Path.GetDirectoryName(mainDocPath.Item2)))
                {
                    System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(mainDocPath.Item2));
                }

                adoc.SaveAs(mainDocPath.Item2);

                foreach (var doc in adoc.ReferencedDocuments)
                {
                    if (paths.Any(x => x.Item1.ToUpper() == doc.FileName.ToUpper()))
                    {
                        var path = paths.First(x => x.Item1.ToUpper() == doc.FileName.ToUpper());

                        if (!System.IO.Directory.Exists(System.IO.Path.GetDirectoryName(path.Item2)))
                        {
                            System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(path.Item2));
                        }

                        doc.SaveAs(path.Item2);
                    }
                }

                adoc.Save();

                adoc.Close();

                adoc.Dispose();

                var newDoc = InventorApplication.Open(mainDocPath.Item2, "");

                CopyHelpers.ReplaceReferences(newDoc, paths);



            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }



        }

        private void InventorGenTemplateButton_Click(object sender, RibbonControlEventArgs e)
        {
            ApplicationType = ApplicationTypeEnum.Inventor;

            AddTemplate();
        }

        #endregion

        #region Solidworks Methods

        private void solidWorksBuildButton_Click(object sender, RibbonControlEventArgs e)
        {
            _processSolidworks = new ProcessRunBlockSolidworks(Globals.ThisAddIn.Application.ActiveSheet);

            try
            {
                _processSolidworks.Run(true);

                if (LogManager.HasData)
                {
                    LogManager.WriteLogs(Globals.ThisAddIn.Application.ActiveWorkbook);

                    MessageBox.Show("Done", "Application Complete - Please Review Logs", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
                else
                {
                    MessageBox.Show("Done", "Application Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ran into an error while running", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                solidWorksStopBuild.Enabled = false;
            }
        }

        private void solidWorksStopBuild_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void solidworksBuildTemplate_Click(object sender, RibbonControlEventArgs e)
        {
            ApplicationType = ApplicationTypeEnum.Solidworks;

            AddTemplate();
        }

        private void solidworksLoadRefDocs_Click(object sender, RibbonControlEventArgs e)
        {
            if (!SolidworksApplication.Attached)
            {
                SolidworksApplication.Attach();
            }

            var activeDocument = SolidworksApplication.ActiveDocument;

            if (activeDocument == null || !activeDocument.IsAssemblyDoc)
            {
                MessageBox.Show("Please open an assembly document");
            }

            var docs = new List<CopyDocumentItem>();

            docs.Add(new CopyDocumentItem(activeDocument.FullFileName));

            var filesPathsToAvoid = Settings.Default.PathsToAvoid.SettingsToList();

            foreach (var doc in activeDocument.Children(false).GetReferencedDocumentsNames())
            {
                if (filesPathsToAvoid.Any(x => doc.ToUpper().Contains(x.ToUpper()))) continue;

                docs.Add(new CopyDocumentItem(doc));
            }

            if (Globals.ThisAddIn.Application.ActiveWorkbook.WorkSheetExists("Copy Tool"))
            {
                Globals.ThisAddIn.Application.ActiveWorkbook.ActivateSheet("Copy Tool");
            }
            else
            {
                Globals.ThisAddIn.Application.ActiveWorkbook.Sheets.Add();

                Globals.ThisAddIn.Application.ActiveSheet.Name = "Copy Tool";
            }

            Excel.Worksheet workSheet = Globals.ThisAddIn.Application.ActiveSheet;

            workSheet.Cells.Clear();

            workSheet.Name = "Copy Tool";

            workSheet.Range["A1"].Value = "Old Path";
            workSheet.Range["B1"].Value = "New Path";

            var i = 2;

            foreach (var d in docs)
            {
                workSheet.Range[$"A{i}"].Value = d.OldPath;
                workSheet.Range[$"B{i}"].Value = d.NewPath;

                i++;
            }

            var copyTable = workSheet.Range[$"A1:B{i}"];

            workSheet.ListObjects.AddEx(XlListObjectSourceType.xlSrcRange, copyTable, null, XlYesNoGuess.xlYes).Name = "Copy Table";

            var outTable = workSheet.GetListObjects().FirstOrDefault(x => x.Name == "Copy Table");

            var old = outTable.GetListColumns().FirstOrDefault(x => x.Name == "Old Path");
        }

        private void solidWorksCopyButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!SolidworksApplication.Attached)
            {
                SolidworksApplication.Attach();
            }

            var activeDocument = SolidworksApplication.ActiveDocument;

            var valid = activeDocument != null && activeDocument.IsAssemblyDoc;

            if (!valid)
            {
                MessageBox.Show("Please open an assembly document");
            }

            var adoc = activeDocument;

            if (Globals.ThisAddIn.Application.ActiveWorkbook.WorkSheetExists("Copy Tool"))
            {
                Globals.ThisAddIn.Application.ActiveWorkbook.ActivateSheet("Copy Tool");
            }
            else
            {
                return;
            }

            Excel.Worksheet workSheet = Globals.ThisAddIn.Application.ActiveSheet;

            try
            {
                var outTable = workSheet.GetListObjects().FirstOrDefault(x => x.Name == "Copy Table");

                if (outTable == null) throw new Exception("Could not find table Copy Table");

                var oldPath = outTable.GetListColumns().FirstOrDefault(x => x.Name == "Old Path");

                var newPath = outTable.GetListColumns().FirstOrDefault(x => x.Name == "New Path");

                var oldPathList = oldPath.Range.GetFilePaths();

                var newPathList = newPath.Range.GetFilePaths();

                if (oldPathList.Count != newPathList.Count)
                {
                    throw new Exception("Old path list and new path list do not match in qty");
                }

                var paths = new List<Tuple<string, string>>();

                for (var i = 0; i < oldPathList.Count; i++)
                {
                    paths.Add(new Tuple<string, string>(oldPathList[i], newPathList[i]));
                }

                var mainDocPath = paths.FirstOrDefault(x => x.Item1.ToUpper() == adoc.FullFileName.ToUpper());

                if (!System.IO.Directory.Exists(System.IO.Path.GetDirectoryName(mainDocPath.Item2)))
                {
                    System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(mainDocPath.Item2));
                }

                adoc.SaveAs(mainDocPath.Item2);

                adoc.Save();

                adoc.Close();

                adoc.Dispose();

                var newDoc = SolidworksApplication.Open(mainDocPath.Item2, DocumentTypes.ASSEMBLY);

                var children = newDoc.Children();

                // capture the current state of suppression
                children.CaptureSuppressionState();

                children.UnsuppressAll();

                foreach (var doc in children.GetReferencedDocuments())
                {
                    if (paths.Any(x => x.Item1.ToUpper() == doc.FullFileName.ToUpper()))
                    {
                        var path = paths.First(x => x.Item1.ToUpper() == doc.FullFileName.ToUpper());

                        if (!System.IO.Directory.Exists(System.IO.Path.GetDirectoryName(path.Item2)))
                        {
                            System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(path.Item2));
                        }

                        doc.SaveAs(path.Item2);
                    }
                }

                SolidworksCopyHelpers solidworksCopy = new SolidworksCopyHelpers();

                solidworksCopy.References(newDoc, paths);

                children.RestoreSuppressionState();

                //CopyHelpers.ReplaceReferences(newDoc, paths);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }
        }

        private void solidworksSettings_Click(object sender, RibbonControlEventArgs e)
        {
            var solidWorksSettingForm = new SolidworksSettingsForm();

            solidWorksSettingForm.Show();
        }

        #endregion

        #region General Methods

        private void AddTemplate()
        {
            Excel.Worksheet workSheet = null;

            var empty = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).WorkSheetEmpty();

            if (empty)
            {
                workSheet = Globals.ThisAddIn.Application.ActiveSheet;
            }
            else
            {
                InputForm form = new InputForm("Enter new Work Sheet Name");

                form.ShowDialog();

                var workSheetName = form.TextInput;

                if (string.IsNullOrEmpty(workSheetName))
                {
                    return;
                }

                Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add();

                workSheet = Globals.ThisAddIn.Application.ActiveSheet;

                workSheet.Name = workSheetName;
            }

            workSheet.Range["A5"].Value = "Part Number";
            workSheet.Range["B5"].Value = "Description";
            workSheet.Range["C5"].Value = "Command";
            workSheet.Range["D5"].Value = "Name";
            workSheet.Range["E5"].Value = "Parent";
            workSheet.Range["F5"].Value = "Value";
            workSheet.Range["G5"].Value = "Value 2";

            workSheet.Range["C5"].Name = $"{workSheet.Name}Type";

            var topRange = workSheet.Range["A5:H5"];

            topRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Orange);
            topRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
            topRange.Font.Bold = true;

            workSheet.Range["C6"].Value = "TopLevelName";
            workSheet.Range["D6"].Value = "Enter Top document name here";

            AddCommands();

            Globals.ThisAddIn.Application.ActiveWorkbook.ActivateSheet(workSheet.Name);

            workSheet.Range["C7"].AddDropDownList("Commands");
        }

        private void AddCommands()
        {
            Excel.Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;

            var workSheet = workbook.GetWorksheets().FirstOrDefault(x => x.Name == ConstantStrings.Commands);

            if (workSheet == null)
            {
                workbook.Worksheets.Add();

                workSheet = Globals.ThisAddIn.Application.ActiveSheet;

                workSheet.Name = "Commands";
            }

            workSheet.Range["A1"].Value = "Command";
            workSheet.Range["B1"].Value = "Name";
            workSheet.Range["C1"].Value = "Parent";
            workSheet.Range["D1"].Value = "Value";
            workSheet.Range["E1"].Value = "Value 2";
            workSheet.Range["F1"].Value = "Notes";

            var commands = GetCommands().Where(x => x.ApplicationType == ApplicationType 
                                                    || x.ApplicationType == ApplicationTypeEnum.General).ToList();

            var notesRange = workSheet.Range[$"A2:G{2 + commands.Count}"];

            notesRange.Clear();

            var i = 2;

            foreach (var c in commands)
            {
                workSheet.Range[$"A{i}"].Value = c.Command;
                workSheet.Range[$"B{i}"].Value = c.Name;
                workSheet.Range[$"C{i}"].Value = c.Parent;
                workSheet.Range[$"D{i}"].Value = c.Value;
                workSheet.Range[$"E{i}"].Value = c.Value2;
                workSheet.Range[$"F{i}"].Value = c.Notes;
                i++;
            }

            workSheet.CreateNamedRange("Commands", $"A2:A{i}");

            // create suppresion validation
            workSheet.ClearNamedRange("Suppression");
            workSheet.Range[$"A{i + 5}"].Value = "S";
            workSheet.Range[$"A{i + 6}"].Value = "U";
            workSheet.CreateNamedRange("Suppression", $"A{i + 5}:A{i + 6}");
        }

        private List<CommandItem> GetCommands()
        {
            return new List<CommandItem>
            {
                new CommandItem(Commands.Dimension, "Name of the dimension followed by the sketch or feature name example \"Dim1@Sketch1\"", ConstantStrings.ParentText, "Value to set the dimension", applicationType: ApplicationTypeEnum.Solidworks),
                new CommandItem(Commands.Equation, "Name of the equation", ConstantStrings.ParentText, "Value to set the equation", notes: "Units Ul, In, MM, CM, M, or test", applicationType: ApplicationTypeEnum.Solidworks),
                new CommandItem(Commands.Parameter, "Name of Parameter", ConstantStrings.ParentText, "Value to set parameter", "Not Used", "Either UL(Unitless), In, MM, CM, M or text", "", ApplicationTypeEnum.Inventor),
                new CommandItem(Commands.GetParameter, "Name of Parameter", ConstantStrings.ParentText, "Application will set the value of the parameter here", "Not Used", "Either UL(Unitless), In, MM, CM, M or text", "", ApplicationTypeEnum.Inventor),
                new CommandItem(Commands.ComponentActivity, "Name of component in the tree followed by occurrence number", ConstantStrings.ParentText, ConstantStrings.SuppressionText),
                new CommandItem(Commands.ConstraintActivity, "Name of constraint", ConstantStrings.ParentText, ConstantStrings.SuppressionText),
                new CommandItem(Commands.PatternActivity, "Name of pattern", ConstantStrings.ParentText, ConstantStrings.SuppressionText),
                new CommandItem(Commands.PlaceComponent, "File location of document", ConstantStrings.ParentText, "Application will set the occurrence name of the part here"),
                new CommandItem(Commands.Stop, "Stops the application", "", ""),
                new CommandItem(Commands.TopLevelName, "Name of the top level document", "", ""),
                new CommandItem(Commands.FeatureActivity, "Name of the feature", ConstantStrings.ParentText, ConstantStrings.SuppressionText),
                new CommandItem(Commands.DeleteComponent, "Name of component in the tree followed by occurrence number", ConstantStrings.ParentText, ""),
                new CommandItem(Commands.DeleteReferencedDocuments, "Name of the document", ConstantStrings.ParentText, "", "", "", "Deletes all components in the parent document that reference the given document"),
                new CommandItem(Commands.ReplaceComponent, "Name of component in the tree followed by occurrence number", ConstantStrings.ParentText, "File location of document to replace the component with"),
                new CommandItem(Commands.Sub, "Name of the named range to run", "Name of the range to set the parameter in", "Value of the parameter to be set"),
                new CommandItem(Commands.If, "", "", "Boolean value to process the lines inside the block if true then the values will be processed if not then it will be skipped", "", "", "Must be followed by a End if"),
                new CommandItem(Commands.EndIf, "", "", "", "", "", "If the condition is false for the matching if the program will skip down the matching end if"),
                new CommandItem(Commands.Repeat, "", "", "Enter the number of times to repeat", "The current value of the repeat", "", "Must be followed with a end repeat"),
                new CommandItem(Commands.EndRepeat, "", "", "", "", "", "Values in between repeat and end repeat will occur until index matches the count number"),
                new CommandItem(Commands.Comment, "", "", "", "", "", "No cells are processed for user comments"),
                new CommandItem(Commands.SetProperty, "Name of the property", ConstantStrings.ParentText, "Value to set the property"),
                new CommandItem(Commands.GetProperty, "Name of the property", ConstantStrings.ParentText, "Application will set the value of the property here"),
                new CommandItem(Commands.SetLevelOfDetail, "Name of the Level Detail", ConstantStrings.ParentText, "", "", "", "Activates the level detail and creates it if not does not exist", applicationType: ApplicationTypeEnum.Inventor),
                new CommandItem(Commands.SetDesignViewRep, "Name of the Design View Representation", ConstantStrings.ParentText, "", "", "", "Activates the Design View and creates it if not does not exist", applicationType: ApplicationTypeEnum.Inventor),
                new CommandItem(Commands.ComponentVisiblity, "Name of component in the tree followed by occurrence number", ConstantStrings.ParentText, "True or False if the component is visible"),
                new CommandItem(Commands.DocumentReferenceVisiblity, "Name of document in the active assembly", ConstantStrings.ParentText, "True or False if the component is visible", "", "", "Sets the visibllity of the all the occurences that reference this document"),
                new CommandItem(Commands.UpdateDocument, "", "", "", "", "", "Updates and saves the active document"),
                new CommandItem(Commands.ShowConfiguration, "Name of configuration to show", ConstantStrings.ParentText, "", applicationType: ApplicationTypeEnum.Solidworks),
                new CommandItem(Commands.SetComponentConfiguration, "Name of the component to set the configuration on", ConstantStrings.ParentText, "Name of the configuration to set on the component", applicationType: ApplicationTypeEnum.Solidworks),
                new CommandItem(Commands.SetWeldmentConfiguration, "Name of the weldment member to set the configuration on", ConstantStrings.ParentText, "Name of the configuration to set on the weldment feature", applicationType: ApplicationTypeEnum.Solidworks),
                 new CommandItem(Commands.OpenDocument, "Name of Document", "Enter the a source document if you want the application to copy source if the requested document does not exist", "Enter text to find in the source delimited by commas", "Enter text to replace with in the source delimited by commas", "", "The Search and replace test is matched by location so if source path is C:\\Blue\\Green\\Yellow and search contains \"Blue, Green, Yellow \" then replace \"Yellow, Blue, Green \" the new source will be C:\\Yellow\\Blue\\Green"),
            };
        }

        #endregion

        private void solidworksCaptureButton_Click(object sender, RibbonControlEventArgs e)
        {
            var captureForm = new CaptureDesignFormSolidworks
            {
                TopMost = true
            };

            captureForm.Show();
        }

        private void captureInventorModelData_Click(object sender, RibbonControlEventArgs e)
        {
            var captureForm = new CaptureDesignFormInventor
            {
                TopMost = true
            };

            captureForm.Show();
        }

        private void InventorSettingsButton_Click(object sender, RibbonControlEventArgs e)
        {
            var settingsForm = new SettingsForm();

            settingsForm.ShowDialog();
        }

        private void InventorCaptureDrawingData_Click(object sender, RibbonControlEventArgs e)
        {
            if (!InventorApplication.Attached)
            {
                InventorApplication.Attach();
            }

            var captureForm = new DrawingCaptureForm();

            captureForm.ShowDialog();
        }
    }
}
