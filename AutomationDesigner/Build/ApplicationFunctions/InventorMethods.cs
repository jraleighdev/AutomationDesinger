using AutomationDesinger.Enums;
using AutomationDesinger.Helpers;
using AutomationDesinger.Logs;
using InventorWrapper;
using InventorWrapper.CopyTools;
using InventorWrapper.Documents;
using InventorWrapper.General;
using InventorWrapper.Helpers;
using InventorWrapper.IProps;
using InventorWrapper.Parameters;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics.Contracts;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;

namespace AutomationDesinger.Build.ApplicationFunctions
{
    public class InventorMethods
    {
        public void SetParameter(InventorDocument document, string paramName, string value)
        {
            var parameter = document.Parameters.FirstOrDefault(x => x.Name == paramName);

            if (parameter == null)
            {
                LogManager.Add($"Could not find parameter {paramName} in {document.Name}");
                return;
            }

            switch (parameter.UnitType)
            {
                case UnitTypes.Length:
                case UnitTypes.Angular:
                    parameter.Value = UnitManager.UnitsToInventor(ConverterHelpers.ConvertDouble(value), parameter.UnitType);
                    break;
                case UnitTypes.Unitless:
                    parameter.Value = ConverterHelpers.ConvertInt(value);
                    break;
                case UnitTypes.Text:
                    parameter.Value = value;
                    break;
            }
        }

        public string GetParameter(InventorDocument workingDocument, string paramName)
        {
            var parameter = workingDocument.Parameters.FirstOrDefault(x => x.Name == paramName);

            if (parameter == null)
            {
                LogManager.Add($"Could not find parameter {paramName} in {workingDocument.Name}");
                return "Not Found";
            }

            switch (parameter.UnitType)
            {
                case UnitTypes.Length:
                case UnitTypes.Angular:
                    return UnitManager.UnitsFromInventor(parameter.Value, parameter.UnitType).ToString();
                case UnitTypes.Unitless:
                    return parameter.Value.ToString();
                case UnitTypes.Text:
                    return parameter.Value.ToString();
                default:
                    return "";
            }
        }

        public void SetProperty(InventorDocument document, string propertyName, string value)
        {
            var properties = GetIpropertyEnums();

            IpropertyEnum iproperty = IpropertyEnum.Custom;

            foreach (var p in properties)
            {
                if (p.GetDescription().ToUpper() == propertyName.ToUpper())
                {
                    iproperty = p;
                    break;
                }
            }

            if (document == null) throw new Exception("Could not find document to set property on");

            try
            {
                document.Properties.SetPropertyValue(iproperty, value, propertyName);
            }
            catch (Exception ex)
            {
                LogManager.Add(ex.Message);
            }
        }

        public string GetProperty(InventorDocument document, string propertyName)
        {
            var properties = GetIpropertyEnums();

            IpropertyEnum iproperty = IpropertyEnum.Custom;

            foreach (var p in properties)
            {
                if (p.GetDescription().ToUpper() == propertyName.ToUpper())
                {
                    iproperty = p;
                    break;
                }
            }

            if (document == null) throw new Exception("Could not find document to set property on");

            var value = document.Properties.GetPropertyValue(iproperty, propertyName);

            if (string.IsNullOrEmpty(value))
            {
                LogManager.Add($"Could not find property {propertyName} in {document.Name}");
            }

            return value;
        }

        public List<IpropertyEnum> GetIpropertyEnums()
        {
            return Enum.GetValues(typeof(IpropertyEnum)).Cast<IpropertyEnum>().ToList();
        }

        public void Suppression(InventorDocument document, string component, string suppressed, SuppresionType suppresionType)
        {
            var status = suppressed.ToUpper() == "S";

            try
            {
                switch (suppresionType)
                {
                    case SuppresionType.Component:
                        document.GetAssemblyDocument().Components.SetComponentStatus(component, status);
                        break;
                    case SuppresionType.Constraint:
                        document.GetAssemblyDocument().Constraints.SetConstraintStatus(component, status);
                        break;
                    case SuppresionType.Pattern:
                        document.GetAssemblyDocument().Patterns.SuppressPattern(component, status);
                        break;
                    case SuppresionType.Feature:
                        document.SetFeatureStatus(component, status);
                        break;
                }
            }
            catch (Exception ex)
            {
                LogManager.Add(ex.Message);
            }
        }

        public void OpenDocument(string name, string sourceLocation, string textToFind, string textToReplace)
        {
            var replaceList = new List<string>();
            var searchList = new List<string>();
            var paths = new List<Tuple<string, string>>();

            if (textToFind.Contains(","))
            {
                searchList.AddRange(textToFind.Split(','));
            }
            else
            {
                searchList.Add(textToFind);
            }

            if (replaceList.Contains(","))
            {
                replaceList.AddRange(textToReplace.Split(','));
            }
            else
            {
                replaceList.Add(textToReplace);
            }

            if (searchList.Count != replaceList.Count)
            {
                throw new Exception("Search list and replace list don't match");
            }

            if (System.IO.File.Exists(name))
            {
                InventorApplication.Open(name, "", true);
            }
            else
            {
                if (System.IO.File.Exists(sourceLocation))
                {
                    var document = InventorApplication.Open(sourceLocation, "", true);

                    var newMainDocPath = document.FileName;

                    for (var i = 0; i < searchList.Count; i++)
                    {
                        newMainDocPath = newMainDocPath.Replace(searchList[i].Trim(), replaceList[i].Trim());
                    }

                    if (!System.IO.Directory.Exists(System.IO.Path.GetDirectoryName(newMainDocPath)))
                    {
                        System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(newMainDocPath));
                    }

                    document.SaveAs(newMainDocPath);

                    if (document.IsAssemblyDoc)
                    {
                        var adoc = document.GetAssemblyDocument();

                        foreach (var doc in adoc.ReferencedDocuments)
                        {
                            var oldPath = doc.FileName;

                            var newPath = oldPath;

                            for (var i = 0; i < searchList.Count; i++)
                            {
                                newPath = newPath.Replace(searchList[i].Trim(), replaceList[i].Trim());
                            }

                            if (!System.IO.Directory.Exists(System.IO.Path.GetDirectoryName(newPath)))
                            {
                                System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(newPath));
                            }

                            doc.SaveAs(newPath);

                            paths.Add(new Tuple<string, string>(oldPath, newPath));
                        }
                    }

                    document.Save();

                    document.Close();

                    document.Dispose();

                    var newDoc = InventorApplication.Open(newMainDocPath, "");

                    if (newDoc.IsAssemblyDoc)
                    {
                        CopyHelpers.ReplaceReferences(newDoc, paths);
                    }
                }
            }
        }

        public void Delete(InventorDocument document, string component)
        {
            try
            {
                document.GetAssemblyDocument().Components.DeleteComponent(component);
            }
            catch (Exception ex)
            {
                LogManager.Add(ex.Message);
            }
        }

        public void DeleteReferenced(InventorDocument document, string component)
        {
            try
            {
                document.GetAssemblyDocument().Components.DeleteComponents(component);
            }
            catch (Exception ex)
            {
                LogManager.Add(ex.Message);
            }
        }
    }
}
