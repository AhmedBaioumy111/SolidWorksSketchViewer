using System;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorksSketchViewer.Models;

namespace SolidWorksSketchViewer.Helpers
{
    public class SolidWorksService
    {
        private SldWorks swApp;
        private ModelDoc2 swModel;

        // Dictionary to store sketch objects with their parent features for ID lookup
        private Dictionary<ISketch, Feature> sketchFeatureMap = new Dictionary<ISketch, Feature>();

        // Dictionary to store dimensions with unique identifiers
        private Dictionary<Dimension, string> dimensionIdentifierMap = new Dictionary<Dimension, string>();
        private int dimensionCounter = 0;

        public SolidWorksService()
        {
            try
            {
                // Try to connect to an already running instance of SOLIDWORKS
                swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
            }
            catch
            {
                // If no instance is running, create a new one
                swApp = new SldWorks();
                swApp.Visible = true;
            }
        }

        private string GetDimensionTypeFromDimension(Dimension dim)
        {
            try
            {
                // Use parent DisplayDimension to determine dimension type
                foreach (KeyValuePair<Dimension, string> entry in dimensionIdentifierMap)
                {
                    if (entry.Key == dim)
                    {
                        string[] parts = entry.Value.Split('|');
                        if (parts.Length > 1 && !string.IsNullOrEmpty(parts[1]))
                        {
                            int dimTypeVal;
                            if (int.TryParse(parts[1], out dimTypeVal))
                            {
                                return GetDimensionType(dimTypeVal);
                            }
                        }
                    }
                }

                // If we can't determine type from stored data, make a best guess
                bool isAngular = false;
                try
                {
                    // Try to detect angular dimensions by value range
                    double value = dim.Value;
                    isAngular = (Math.Abs(value) <= 6.28); // Rough heuristic for radians
                }
                catch
                {
                    isAngular = false;
                }

                return isAngular ? "Angular" : "Linear";
            }
            catch
            {
                return "Unknown";
            }
        }

        public SolidWorksDocumentModel OpenDocument(string filePath)
        {
            // Clear stored maps when opening a new document
            sketchFeatureMap.Clear();
            dimensionIdentifierMap.Clear();
            dimensionCounter = 0;

            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            {
                return null;
            }

            // Open the document
            int errors = 0;
            int warnings = 0;
            swModel = swApp.OpenDoc6(filePath, (int)swDocumentTypes_e.swDocPART,
                                     (int)swOpenDocOptions_e.swOpenDocOptions_ReadOnly,
                                     "", ref errors, ref warnings);

            if (swModel == null)
            {
                return null;
            }

            // Create a document model
            var documentModel = new SolidWorksDocumentModel
            {
                FilePath = filePath,
                FileName = Path.GetFileName(filePath)
            };

            // Set document type
            switch (swModel.GetType())
            {
                case (int)swDocumentTypes_e.swDocPART:
                    documentModel.DocumentType = "Part";
                    break;
                case (int)swDocumentTypes_e.swDocASSEMBLY:
                    documentModel.DocumentType = "Assembly";
                    break;
                case (int)swDocumentTypes_e.swDocDRAWING:
                    documentModel.DocumentType = "Drawing";
                    break;
                default:
                    documentModel.DocumentType = "Unknown";
                    break;
            }

            // Populate sketches
            documentModel.Sketches = GetAllSketches();

            return documentModel;
        }

        public ObservableCollection<SketchModel> GetAllSketches()
        {
            var sketches = new ObservableCollection<SketchModel>();

            if (swModel == null)
            {
                return sketches;
            }

            // Get the feature manager
            FeatureManager featureManager = swModel.FeatureManager;
            object featureObj = swModel.FirstFeature();

            if (featureObj == null)
                return sketches;

            Feature feature = null;
            try
            {
                feature = (Feature)featureObj;
            }
            catch
            {
                return sketches;
            }

            // Loop through all features
            while (feature != null)
            {
                if (feature.GetTypeName2() == "ProfileFeature" || feature.GetTypeName2() == "SketchFeature")
                {
                    var sketch = new SketchModel
                    {
                        Name = feature.Name,
                        Id = feature.GetID().ToString(),
                        FeatureType = feature.GetTypeName2(),
                        IsVisible = GetSketchVisibility(feature)
                    };

                    // Get the sketch entities count
                    object sketchObjRaw = feature.GetSpecificFeature2();
                    if (sketchObjRaw != null)
                    {
                        Sketch sketchObject = null;
                        try
                        {
                            sketchObject = (Sketch)sketchObjRaw;
                        }
                        catch
                        {
                            // Skip if casting fails
                            continue;
                        }

                        // Store sketch with its parent feature for later ID lookup
                        if (!sketchFeatureMap.ContainsKey(sketchObject))
                        {
                            sketchFeatureMap.Add(sketchObject, feature);
                        }

                        // Get entity count by checking segment count
                        object segmentsObj = sketchObject.GetSketchSegments();
                        object pointsObj = sketchObject.GetSketchPoints2();

                        object[] segments = segmentsObj as object[];
                        object[] points = pointsObj as object[];

                        sketch.EntityCount = 0;
                        if (segments != null)
                            sketch.EntityCount += segments.Length;
                        if (points != null)
                            sketch.EntityCount += points.Length;

                        // Check if sketch is active
                        ISketch activeSketch = swModel.SketchManager.ActiveSketch;
                        sketch.IsActive = (activeSketch != null && activeSketch == sketchObject);
                    }

                    // Get dimensions for the sketch
                    sketch.Dimensions = GetSketchDimensions(feature);

                    sketches.Add(sketch);
                }

                // Get the next feature
                object nextFeatureObj = feature.GetNextFeature();
                if (nextFeatureObj == null)
                    break;

                try
                {
                    feature = (Feature)nextFeatureObj;
                }
                catch
                {
                    break;
                }
            }

            return sketches;
        }

        public ObservableCollection<DimensionModel> GetSketchDimensions(Feature sketchFeature)
        {
            var dimensions = new ObservableCollection<DimensionModel>();

            if (sketchFeature == null)
            {
                return dimensions;
            }

            // Get the first display dimension
            object dispDimObj = sketchFeature.GetFirstDisplayDimension();

            if (dispDimObj == null)
                return dimensions;

            DisplayDimension dispDim = null;
            try
            {
                dispDim = (DisplayDimension)dispDimObj;
            }
            catch
            {
                return dimensions;
            }

            // Loop through all dimensions using the iterator pattern
            while (dispDim != null)
            {
                object dimObj = dispDim.GetDimension();
                if (dimObj != null)
                {
                    Dimension dim = null;
                    try
                    {
                        dim = (Dimension)dimObj;

                        // Store dimension with type info for later lookup
                        // Get type from DisplayDimension and store it
                        int dispDimType = dispDim.GetType();
                        string dimIdentifier = dimensionCounter + "|" + dispDimType;
                        dimensionIdentifierMap[dim] = dimIdentifier;
                        dimensionCounter++;

                        var dimensionModel = new DimensionModel
                        {
                            Name = dim.GetNameForSelection(),
                            Value = dim.Value.ToString(), // Get value in active config
                            Type = GetDimensionType(dispDimType),
                            IsReference = dim.IsReference(),
                            Units = GetDimensionUnits(dispDim)
                        };

                        dimensions.Add(dimensionModel);
                    }
                    catch
                    {
                        // Skip if casting fails
                    }
                }

                // Get next dimension
                object nextDispDimObj = sketchFeature.GetNextDisplayDimension(dispDim);
                if (nextDispDimObj == null)
                    break;

                try
                {
                    dispDim = (DisplayDimension)nextDispDimObj;
                }
                catch
                {
                    break;
                }
            }

            return dimensions;
        }

        private bool GetSketchVisibility(Feature feature)
        {
            try
            {
                // Check if the feature is suppressed using the IsSuppressed method
                return !feature.IsSuppressed();
            }
            catch
            {
                return false;
            }
        }

        private string GetDimensionType(int dimType)
        {
            switch (dimType)
            {
                case 0: // swAngularDimension
                    return "Angular";
                case 1: // swArcLengthDimension
                    return "Arc Length";
                case 2: // swChamferDimension  
                    return "Chamfer";
                case 3: // swDiameterDimension
                    return "Diameter";
                case 4: // swLinearDimension
                    return "Linear";
                case 5: // swOrdinateDimension
                    return "Ordinate";
                case 6: // swRadialDimension
                    return "Radial";
                default:
                    return "Unknown";
            }
        }

        private string GetDimensionUnits(DisplayDimension dispDim)
        {
            try
            {
                if (dispDim == null)
                    return "Unknown";

                // Check if using document units
                if (dispDim.GetUseDocUnits())
                    return "Document Units";

                // Get the units from the display dimension
                int units = dispDim.GetUnits();

                // Get the dimension to check its type
                object dimObj = dispDim.GetDimension();
                if (dimObj == null)
                    return "Unknown";

                Dimension dim = null;
                try
                {
                    dim = (Dimension)dimObj;
                }
                catch
                {
                    return "Unknown";
                }

                // Get dimension type from display dimension
                int dimType = dispDim.GetType();
                bool isAngular = (dimType == 0); // 0 = swAngularDimension

                // Check if angular or linear
                if (isAngular)
                {
                    // For angular dimensions
                    switch (units)
                    {
                        case 0: // swDegrees
                            return "deg";
                        case 1: // swRadians
                            return "rad";
                        default:
                            return "Unknown";
                    }
                }
                else
                {
                    // For linear dimensions
                    switch (units)
                    {
                        case 0: // swMM
                            return "mm";
                        case 1: // swCM
                            return "cm";
                        case 2: // swMETER
                            return "m";
                        case 3: // swINCHES
                            return "in";
                        case 4: // swFEET
                            return "ft";
                        default:
                            return "Unknown";
                    }
                }
            }
            catch
            {
                return "Unknown";
            }
        }

        public void SelectSketch(string sketchId)
        {
            if (swModel == null || string.IsNullOrEmpty(sketchId))
            {
                return;
            }

            // Try to get the feature by ID
            Feature feature = null;
            try
            {
                // Use alternative method to find feature by ID
                feature = FindFeatureById(swModel, sketchId);
            }
            catch
            {
                // Silently catch errors
            }

            if (feature != null)
            {
                // Clear current selection
                swModel.ClearSelection2(true);

                // Select the feature
                feature.Select2(false, 0);

                // Optionally, zoom to fit
                swModel.ViewZoomtofit2();
            }
        }

        // Helper method to find feature by ID
        private Feature FindFeatureById(ModelDoc2 model, string featureId)
        {
            if (model == null)
                return null;

            object featObj = model.FirstFeature();
            if (featObj == null)
                return null;

            Feature feat = null;
            try
            {
                feat = (Feature)featObj;
            }
            catch
            {
                return null;
            }

            while (feat != null)
            {
                if (feat.GetID().ToString() == featureId)
                    return feat;

                // Check child features
                object subFeatObj = feat.GetFirstSubFeature();
                if (subFeatObj != null)
                {
                    Feature subFeat = null;
                    try
                    {
                        subFeat = (Feature)subFeatObj;
                    }
                    catch
                    {
                        subFeat = null;
                    }

                    while (subFeat != null)
                    {
                        if (subFeat.GetID().ToString() == featureId)
                            return subFeat;

                        object nextSubFeatObj = subFeat.GetNextSubFeature();
                        if (nextSubFeatObj == null)
                            break;

                        try
                        {
                            subFeat = (Feature)nextSubFeatObj;
                        }
                        catch
                        {
                            break;
                        }
                    }
                }

                object nextFeatObj = feat.GetNextFeature();
                if (nextFeatObj == null)
                    break;

                try
                {
                    feat = (Feature)nextFeatObj;
                }
                catch
                {
                    break;
                }
            }

            return null;
        }

        public void CloseDocument()
        {
            if (swModel != null)
            {
                swApp.CloseDoc(swModel.GetTitle());
                swModel = null;
            }
        }

        public void Dispose()
        {
            CloseDocument();
            if (swApp != null)
            {
                // Don't terminate SOLIDWORKS if we didn't start it
                // swApp.ExitApp(); 
                swApp = null;
            }

            // Clear dictionaries
            sketchFeatureMap.Clear();
            dimensionIdentifierMap.Clear();
        }
    }
}