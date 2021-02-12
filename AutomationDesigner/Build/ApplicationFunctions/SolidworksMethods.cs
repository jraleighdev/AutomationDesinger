using AutomationDesinger.Enums;
using AutomationDesinger.Logs;
using SolidworksWrapper;
using SolidworksWrapper.Constants;
using SolidworksWrapper.Documents;
using SolidworksWrapper.Enums;
using SolidworksWrapper.General;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutomationDesinger.Build.ApplicationFunctions
{
    public class SolidworksMethods
    {
        public void Test()
        {
            var doc = SolidworksApplication.ActiveDocument;

            var dim = doc.GetSolidworksDimension("D1@Sketch1");
        }

        public void SetDimValue(SolidworksDocument workingDocument, string name, string configuration, double value)
        {
            var dim = workingDocument.GetSolidworksDimension(name);

            if (dim == null)
            {
                LogManager.Add($"Could not find dimension {name} in {workingDocument.Name}");
                return;
            }

            double convertedValue;

            switch (dim.Type)
            {
                case DimensionTypes.Linear:
                    convertedValue = UnitManager.UnitsToSolidworks(value);
                    break;
                case DimensionTypes.Angular:
                    convertedValue = UnitManager.ConvertDegrees(value);
                    break;
                case DimensionTypes.Integer:
                default:
                    convertedValue = value;
                    break;  
            }

            var success = dim.SetValue(convertedValue, string.IsNullOrEmpty(configuration) ? ConfigurationOptions.ThisConfiguration
                                                                                           : ConfigurationOptions.SpecifyConfiguration,
                                                                                           configuration);
            if (!success)
            {
                LogManager.Add($"Could not set dim {name} in {workingDocument.Name}");
            }
        }

        public string GetDimValue(SolidworksDocument workingDocument, string name, string configuration)
        {
            var dim = workingDocument.GetSolidworksDimension(name);

            if (dim == null)
            {
                LogManager.Add($"Could not find dimension {name} in {workingDocument.Name}");
                return "";
            }

            double value = dim.GetValue(string.IsNullOrEmpty(configuration) ? ConfigurationOptions.ThisConfiguration 
                                                                            : ConfigurationOptions.SpecifyConfiguration, 
                                                                            configuration);
            double convertedValue;
            switch (dim.Type)
            {
                case DimensionTypes.Linear:
                    convertedValue = UnitManager.UnitsFromSolidworks(value);
                    break;
                case DimensionTypes.Angular:
                    convertedValue = UnitManager.ConvertRadians(value);
                    break;
                case DimensionTypes.Integer:
                default:
                    convertedValue = value;
                    break;
            }

            return convertedValue.ToString();
        }

        public void SetEquation(SolidworksDocument workingDocument, string name, double value)
        {
            workingDocument.Equations.SetEquation(name, value);
        }

        public string GetEquation(SolidworksDocument workingDocument, string name)
        {
            var equation = workingDocument.Equations.GetEquation(name);

            if (equation == null)
            {
                LogManager.Add($"Could not find equation {name} in {workingDocument.Name}");
            }

            return equation == null ? "" : equation.Value.ToString();
        }

        public void SetVisiblity(SolidworksDocument document, string name, bool visibilty, string featureType)
        {
            document.ClearSelection();

            try
            {
                switch (featureType)
                {
                    case FeatureTypes.Component:
                        document.Select(name, FeatureTypes.Component);
                        break;
                    
                }

                if (document.SelectedCount() != 1)
                {
                    throw new Exception($"Could not find {name} in {document.Name}");
                }

                if (visibilty)
                {
                    document.ShowSelected();
                }
                else
                {
                    document.HideSelected();
                }
            }
            catch (Exception ex)
            {
                LogManager.Add(ex.Message);
            }
        }

        public void Suppression(SolidworksDocument document, string name, string value, SuppresionType suppresionType)
        {
            var status = value.ToUpper() == "S";

            document.ClearSelection();

            try
            {
                switch (suppresionType)
                {
                    case SuppresionType.Component:
                        document.Select(name, FeatureTypes.Component);
                        break;
                    case SuppresionType.Constraint:
                        document.Select(name, FeatureTypes.Mate);
                        break;
                    case SuppresionType.Feature:
                        document.Select(name, FeatureTypes.BodyFeature);
                        break;
                    case SuppresionType.Pattern:
                        document.Select(name, FeatureTypes.ComponentPattern);
                        break;
                    case SuppresionType.Folder:
                        document.Select(name, FeatureTypes.Folder);
                        break;
                }

                if (document.SelectedCount() != 1)
                {
                    throw new Exception($"Could not find {name} in {document.Name}");
                }

                if (status)
                {
                    document.SuppressSelected();
                }
                else
                {
                    document.UnsuppressSelected();
                }

            }
            catch (Exception ex)
            {
                LogManager.Add(ex.Message);
            }
        }

        public void SetComponentConfiguration(SolidworksDocument document, string componentName, string configurationName)
        {
            document.ClearSelection();

            document.Select(componentName, FeatureTypes.Component);

            if (document.SelectedCount() != 1)
            {
                LogManager.Add($"Could not find {componentName} in {document.Name}");
                return;
            }

            var comp = document.GetSelectedComponent();

            if (comp != null)
            {
                comp.ReferencedConfiguration = configurationName;
            }
        }

        public void SetProperty(SolidworksDocument document, string propertyName, string value)
        {
            try
            {
                document.ActiveConfiguration.Properties.SetValue(propertyName, value);
            }
            catch (Exception ex)
            {
                LogManager.Add(ex.Message);
            }
        }

        public  void SetDocumentReferenceVisibility(SolidworksDocument workingDocument, string name, bool visible, string featureType)
        {
            if (workingDocument.IsAssemblyDoc)
            {
                workingDocument.ClearSelection();

                var children = workingDocument.Children().Where(c => c.SolidworksDocument.Name == name);

                foreach (var c in children)
                {
                    workingDocument.Select(c.Name, FeatureTypes.Component);

                    if (workingDocument.SelectedCount() == 1)
                    {
                        if (visible)
                        {
                            workingDocument.ShowSelected();
                        }
                        else
                        {
                            workingDocument.HideSelected();
                        }
                    }

                    workingDocument.ClearSelection();
                }
            }
        }

        public string GetProperty(SolidworksDocument document, string propertyName)
        {
            var value = "";

            try
            {
                value = document.ActiveConfiguration.Properties.GetValue(propertyName);
            }
            catch (Exception ex)
            {
                LogManager.Add(ex.Message);
            }

            return value;
        }

        public void ShowConfiguration(SolidworksDocument document, string configurationName)
        {
            if (document.ConfigurationExists(configurationName))
            {
                document.ShowConfiguration(configurationName);
            }
            else
            {
                LogManager.Add($"{configurationName} does not exist in document {document.Name}");
            }
        }

        public void SetWeldmentMemberConfiguration(SolidworksDocument document, string featurename, string configurationName)
        {
            document.ClearSelection();

            document.Select(featurename, FeatureTypes.BodyFeature);

            if (document.SelectedCount() != 1)
            {
                LogManager.Add($"Could not find {featurename} in {document.Name}");
                return;
            }

            var feature = document.GetSelectedFeature();

            if (feature.TypeName.ToUpper() == FeatureSubTypes.WeldMemberFeat.ToUpper())
            {
                feature.SetWeldmentConfiguration(configurationName);
            }
            else
            {
                LogManager.Add($"Feature {featurename} in {document.Name} is not a weldment member");
                return;
            }

        }
    }

}
