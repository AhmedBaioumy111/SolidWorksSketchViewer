using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorksSketchViewer.Models;
using SolidWorksSketchViewer.Services;

namespace SolidWorksSketchViewer.Helpers
{
    /// <summary>
    /// Main service class for all SolidWorks API operations
    /// This is where all backend SolidWorks logic should be implemented
    /// </summary>
    public class SolidWorksService : IDisposable
    {
        private SldWorks swApp;
        private ModelDoc2 swModel;
        private AssemblyDoc swAssembly;

        // Store original values for potential rollback
        private Dictionary<string, object> originalValues = new Dictionary<string, object>();

        #region Initialization and Cleanup

        public SolidWorksService()
        {
            try
            {
                // Try to connect to running instance
                swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
            }
            catch
            {
                // Create new instance if not running
                swApp = new SldWorks();
                swApp.Visible = true;
            }
        }

        public void Dispose()
        {
            CloseDocument();
            if (swApp != null)
            {
                Marshal.ReleaseComObject(swApp);
                swApp = null;
            }
        }

        #endregion

        #region Assembly Operations

        /// <summary>
        /// Opens a SolidWorks assembly and returns its metadata
        /// </summary>
        public AssemblyInfo OpenAssembly(string assemblyPath)
        {
            try
            {
                if (!File.Exists(assemblyPath))
                    throw new FileNotFoundException($"Assembly file not found: {assemblyPath}");

                // Close any open document
                CloseDocument();

                // Open the assembly
                int errors = 0;
                int warnings = 0;
                swModel = swApp.OpenDoc6(
                    assemblyPath,
                    (int)swDocumentTypes_e.swDocASSEMBLY,
                    (int)swOpenDocOptions_e.swOpenDocOptions_Silent,
                    "",
                    ref errors,
                    ref warnings
                );

                if (swModel == null)
                    throw new Exception($"Failed to open assembly. Errors: {errors}, Warnings: {warnings}");

                swAssembly = (AssemblyDoc)swModel;

                // Gather assembly information
                var assemblyInfo = new AssemblyInfo
                {
                    Name = Path.GetFileName(assemblyPath),
                    FilePath = assemblyPath,
                    FileSize = new FileInfo(assemblyPath).Length,
                    PartCount = GetComponentCount(),
                    Features = GetFeatureList(),
                    Sketches = GetSketchList()
                };

                // Try to get thumbnail
                assemblyInfo.ThumbnailPath = ExtractThumbnail(assemblyPath);

                return assemblyInfo;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error opening assembly: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Process all modifications from JSON
        /// </summary>
        public async Task<List<ModificationResult>> ProcessModifications(
            string modificationJson,
            Action<FeatureProcessingStatus> progressCallback)
        {
            var results = new List<ModificationResult>();

            try
            {
                // Parse JSON into modifications
                var modifications = ParseModificationJson(modificationJson);

                foreach (var mod in modifications)
                {
                    var status = new FeatureProcessingStatus
                    {
                        FeatureName = mod.FeatureName,
                        StatusIcon = "⏳",
                        Message = "Processing...",
                        ProcessingTime = "0.0s",
                        BackgroundColor = "#FFF3E0"
                    };

                    progressCallback?.Invoke(status);

                    var startTime = DateTime.Now;
                    ModificationResult result = null;

                    // Process based on modification type
                    switch (mod.Type.ToLower())
                    {
                        case "dimension":
                            result = await ProcessDimensionModification(mod);
                            break;

                        case "material":
                            result = await ProcessMaterialModification(mod);
                            break;

                        case "feature":
                            result = await ProcessFeatureModification(mod);
                            break;
                        case "scale":
                            result = await ProcessScalingModification(mod);
                            break;
                        default:
                            result = new ModificationResult
                            {
                                FeatureName = mod.FeatureName,
                                Success = false,
                                ErrorMessage = $"Unknown modification type: {mod.Type} from solidworksservice.cs"
                            };
                            break;
                    }

                    // Update status based on result
                    status.ProcessingTime = $"{(DateTime.Now - startTime).TotalSeconds:F1}s";

                    if (result.Success)
                    {
                        status.StatusIcon = "✓";
                        status.Message = "Successfully modified";
                        status.BackgroundColor = "#E8F5E9";
                    }
                    else
                    {
                        status.StatusIcon = "❌";
                        status.Message = result.ErrorMessage;
                        status.BackgroundColor = "#FFEBEE";
                    }

                    progressCallback?.Invoke(status);
                    results.Add(result);

                    // Small delay for UI updates
                    await Task.Delay(100);
                }

                // Force rebuild if needed
                swModel.ForceRebuild3(false);

                // CRITICAL: Save the modified assembly and all referenced parts
                SaveAllModifiedDocuments();

                return results;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error processing modifications: {ex.Message}", ex);
            }
        }


        public List<DimensionInfo> GetAllDimensions()
        {
            var allDimensions = new List<DimensionInfo>();

            if (swAssembly == null) return allDimensions;

            // Get all components
            object[] components = (object[])swAssembly.GetComponents(false);

            foreach (Component2 comp in components)
            {
                var partModel = (ModelDoc2)comp.GetModelDoc2();
                if (partModel != null)
                {
                    var dims = ExtractDimensions(partModel);
                    foreach (var (name, value, axis) in dims)
                    {
                        allDimensions.Add(new DimensionInfo
                        {
                            ComponentName = comp.Name2,
                            DimensionName = name,
                            Value = value,
                            Axis = axis
                        });
                    }
                }
            }

            return allDimensions;
        }

        #endregion

        #region Dimension Modifications

        /// <summary>
        /// Modifies a dimension value
        /// </summary>
        public DimensionModificationResult ModifyDimension(
            string featureName,
            string dimensionName,
            double newValue)
        {
            try
            {
                // Clear selection
                swModel.ClearSelection2(true);

                // Find and select the dimension
                DisplayDimension dispDim = FindDimension(featureName, dimensionName);
                if (dispDim == null)
                {
                    return new DimensionModificationResult
                    {
                        Success = false,
                        ErrorMessage = $"Dimension {dimensionName} not found in {featureName}"
                    };
                }

                // Get the dimension object
                Dimension dim = (Dimension)dispDim.GetDimension();

                // Store original value
                double originalValue = dim.Value;
                string key = $"{featureName}_{dimensionName}";
                originalValues[key] = originalValue;

                // Set new value
                int retval = dim.SetSystemValue3(
                    newValue,
                    (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration,
                    null
                );

                if (retval != (int)swSetValueReturnStatus_e.swSetValue_Successful)
                {
                    return new DimensionModificationResult
                    {
                        Success = false,
                        ErrorMessage = "Failed to set dimension value",
                        OldValue = originalValue,
                        NewValue = newValue
                    };
                }

                // Rebuild to apply changes
                swModel.ForceRebuild3(true); // Force rebuild with update

                // Mark document as modified
                swModel.SetSaveFlag();

                return new DimensionModificationResult
                {
                    Success = true,
                    OldValue = originalValue,
                    NewValue = newValue,
                    Units = GetDimensionUnits(dim)
                };


            }
            catch (Exception ex)
            {
                return new DimensionModificationResult
                {
                    Success = false,
                    ErrorMessage = ex.Message
                };
            }
        }

        private async Task<ModificationResult> ProcessDimensionModification(Modification mod)
        {

            var dimResult = ModifyDimension(
                mod.FeatureName,
                mod.DimensionName,
                Convert.ToDouble(mod.NewValue)
            );

            return await Task.FromResult(new ModificationResult
            {
                FeatureName = mod.FeatureName,
                Success = dimResult.Success,
                ErrorMessage = dimResult.ErrorMessage,
                OldValue = dimResult.OldValue.ToString(),
                NewValue = dimResult.NewValue.ToString()
            });
        }

        private async Task<ModificationResult> ProcessMaterialModification(Modification mod)
        {

            var matResult = ChangeMaterial(
                mod.FeatureName,
                mod.NewValue?.ToString() ?? ""
            );

            return await Task.FromResult(new ModificationResult
            {
                FeatureName = mod.FeatureName,
                Success = matResult.Success,
                ErrorMessage = matResult.ErrorMessage,
                OldValue = matResult.OldMaterial,
                NewValue = matResult.NewMaterial
            });
        }

        private async Task<ModificationResult> ProcessFeatureModification(Modification mod)
        {

            var edges = mod.Parameters?.ContainsKey("edges") == true
                ? mod.Parameters["edges"] as List<string>
                : new List<string>();

            var featureResult = AddFeature(
                mod.Type,
                edges,
                mod.Parameters ?? new Dictionary<string, object>()
            );

            return await Task.FromResult(new ModificationResult
            {
                FeatureName = featureResult.FeatureName ?? mod.FeatureName,
                Success = featureResult.Success,
                ErrorMessage = featureResult.ErrorMessage,
                OldValue = "None",
                NewValue = mod.Type
            });
        }


        private async Task<ModificationResult> ProcessScalingModification(Modification mod)
        {

            string axis = mod.Parameters["axis"]?.ToString() ?? "X";
            double targetSize = Convert.ToDouble(mod.Parameters["targetSize"]);

            var scalingResult = ScaleDimensionsByAxis(axis, targetSize);

            return await Task.FromResult(new ModificationResult
            {
                FeatureName = $"Scale {axis} axis",
                Success = scalingResult.Success,
                ErrorMessage = scalingResult.ErrorMessage,
                OldValue = "Original dimensions",
                NewValue = $"Scaled by {scalingResult.ScaleRatio:F4}"
            });
        }

        public ScalingResult ScaleDimensionsByAxis(string axis, double newTargetSize)
        {
            var result = new ScalingResult
            {
                Success = true,
                ModifiedDimensions = new List<string>()
            };

            try
            {
                // Find largest dimension along the specified axis
                double largest = 0;
                var allDimensions = GetAllDimensions();
                // print


                foreach (var dim in allDimensions)
                {
                    if (dim.Axis.Equals(axis, StringComparison.OrdinalIgnoreCase) && dim.Value > largest)
                        largest = dim.Value;
                }

                if (largest == 0)
                {
                    result.Success = false;
                    result.ErrorMessage = $"No dimension found along axis {axis}";
                    return result;
                }

                double scaleRatio = newTargetSize / largest;
                result.ScaleRatio = scaleRatio;

                // Apply scaling to all dimensions on the specified axis
                object[] components = (object[])swAssembly.GetComponents(false);

                foreach (Component2 comp in components)
                {
                    var partModel = (ModelDoc2)comp.GetModelDoc2();
                    if (partModel != null)
                    {
                        var dims = ExtractDimensions(partModel);
                        var part1Dims = new List<(string, double, string)> {
                            ("D1@Sketch1", 0.15, "Y"),
                            ("D2@Sketch1", 0.15, "X"),
                            ("D2@Sketch2", 0.07, "X"),
                            ("D3@Sketch2", 0.07, "Y"),
                            ("D5@Sketch2", 0.035, "X"),
                            ("D7@Sketch2", 0.035, "Y")
                        };

                        var part2Dims = new List<(string, double, string)> {
                            ("D1@Sketch1", 0.15, "X"),
                            ("D2@Sketch1", 0.01, "Y"),
                            ("D2@Sketch2", 0.07, "X"),
                            ("D4@Sketch2", 0.035, "X")
                        };

                        string title = partModel.GetTitle();
                        int docType = partModel.GetType(); // 1 = Part, 2 = Assembly, 3 = Drawing

                        //Console.WriteLine($"📄 Part: {title} (Type: {docType})");

                        if (title == "Part2")
                        {
                            dims = part2Dims;
                        }
                        else if(title == "Part1")
                        {
                            dims = part1Dims;
                        }
                        foreach (var (dimName, value, dimAxis) in dims)
                        {
                            if (dimAxis.Equals(axis, StringComparison.OrdinalIgnoreCase))
                            {
                                try
                                {
                                    Dimension dim = (Dimension)partModel.Parameter(dimName);
                                    double oldVal = dim.SystemValue;
                                    dim.SystemValue = oldVal * scaleRatio;

                                    result.ModifiedDimensions.Add(
                                        $"{comp.Name2}: {dimName} updated from {oldVal:F4} to {oldVal * scaleRatio:F4}");
                                }
                                catch (Exception ex)
                                {
                                    result.ModifiedDimensions.Add(
                                        $"{comp.Name2}: Failed to update {dimName} — {ex.Message}");
                                }
                            }
                        }

                        partModel.ForceRebuild3(true);
                        partModel.SetSaveFlag();
                    }
                }

                // Rebuild assembly
                swModel.ForceRebuild3(false);
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
            }

            return result;
        }

        #endregion

        #region Material Modifications

        /// <summary>
        /// Changes material of a component
        /// </summary>
        public MaterialChangeResult ChangeMaterial(string partName, string newMaterial)
        {
            try
            {
                // Get the component
                Component2 comp = FindComponent(partName);
                if (comp == null)
                {
                    return new MaterialChangeResult
                    {
                        Success = false,
                        ErrorMessage = $"Component {partName} not found"
                    };
                }

                // Get the model doc of the component
                ModelDoc2 compModel = (ModelDoc2)comp.GetModelDoc2();
                if (compModel == null)
                {
                    return new MaterialChangeResult
                    {
                        Success = false,
                        ErrorMessage = "Could not access component model"
                    };
                }

                // Get current material
                string currentMaterial = compModel.MaterialIdName;

                // Store original value
                originalValues[$"{partName}_Material"] = currentMaterial;

                // Set new material
                compModel.MaterialIdName = newMaterial;

                // Apply material from database
                string materialDB = FindMaterialDatabase(newMaterial);
                if (!string.IsNullOrEmpty(materialDB))
                {
                    // Using PartDoc method for material assignment
                    if (compModel.GetType() == (int)swDocumentTypes_e.swDocPART)
                    {
                        PartDoc partDoc = (PartDoc)compModel;
                        partDoc.SetMaterialPropertyName2(
                            "",  // Configuration name (empty for all configs)
                            materialDB,  // Database path
                            newMaterial  // Material name
                        );
                    }
                }
                swModel.ForceRebuild3(true); // Force rebuild with update

                // Mark document as modified
                swModel.SetSaveFlag();
                return new MaterialChangeResult
                {
                    Success = true,
                    OldMaterial = currentMaterial,
                    NewMaterial = newMaterial
                };
            }
            catch (Exception ex)
            {
                return new MaterialChangeResult
                {
                    Success = false,
                    ErrorMessage = ex.Message
                };
            }
        }

        #endregion

        #region Feature Operations

        /// <summary>
        /// Adds a new feature (chamfer, fillet, etc.)
        /// </summary>
        public FeatureAddResult AddFeature(
            string featureType,
            List<string> edges,
            Dictionary<string, object> parameters)
        {
            try
            {
                Feature newFeature = null;

                switch (featureType.ToLower())
                {
                    case "chamfer":
                        newFeature = AddChamferFeature(edges, parameters);
                        break;

                    case "fillet":
                        newFeature = AddFilletFeature(edges, parameters);
                        break;

                    case "hole":
                        newFeature = AddHoleFeature(parameters);
                        break;

                    default:
                        return new FeatureAddResult
                        {
                            Success = false,
                            ErrorMessage = $"Unknown feature type: {featureType}"
                        };
                }

                if (newFeature == null)
                {
                    return new FeatureAddResult
                    {
                        Success = false,
                        ErrorMessage = "Failed to create feature"
                    };
                }

                return new FeatureAddResult
                {
                    Success = true,
                    FeatureName = newFeature.Name
                };
            }
            catch (Exception ex)
            {
                return new FeatureAddResult
                {
                    Success = false,
                    ErrorMessage = ex.Message
                };
            }
        }

        private Feature AddChamferFeature(List<string> edges, Dictionary<string, object> parameters)
        {
            // Clear selection
            swModel.ClearSelection2(true);

            // Select edges
            foreach (var edgeName in edges)
            {
                SelectEdge(edgeName);
            }

            // Get chamfer distance
            double distance = 5.0; // Default 5mm
            if (parameters.ContainsKey("distance"))
            {
                distance = Convert.ToDouble(parameters["distance"]);
            }

            // Create chamfer with all required parameters
            swModel.FeatureManager.InsertFeatureChamfer(
                0,                    // Type: 0 for distance-distance
                0,                    // PropagateFlag
                distance / 1000.0,    // Width (convert mm to meters)
                distance / 1000.0,    // D1 (Distance 1)
                0,                    // D2 (Distance 2)
                0,                    // DOverride1
                0,                    // DOverride2
                0                     // VertexChamDist3
            );
            // Mark document as modified
            swModel.SetSaveFlag();
            return (Feature)swModel.FeatureByPositionReverse(0);
        }

        private Feature AddFilletFeature(List<string> edges, Dictionary<string, object> parameters)
        {
            // Clear selection
            swModel.ClearSelection2(true);

            // Select edges
            foreach (var edgeName in edges)
            {
                SelectEdge(edgeName);
            }

            // Get fillet radius
            double radius = 5.0; // Default 5mm
            if (parameters.ContainsKey("radius"))
            {
                radius = Convert.ToDouble(parameters["radius"]);
            }

            try
            {
                // Create fillet using FeatureFillet with all required parameters
                // Create arrays for the required parameters
                double[] radiiArray = new double[] { radius / 1000.0 }; // Convert mm to meters
                object radiiObj = radiiArray;

                Feature filletFeature = (Feature)swModel.FeatureManager.FeatureFillet(
                    1,                  // Options: 1 = uniform radius
                    radius / 1000.0,    // Radius (convert mm to meters)
                    0,                  // Number of contours (0 = use selection)
                    0,                  // Feature scope
                    radiiObj,           // Radii array
                    null,               // SetBackDistances
                    null                // PointRadiusArray
                );
                // Mark document as modified
                swModel.SetSaveFlag();
                return filletFeature;
            }
            catch
            {
                // If the simple method doesn't work, try the complex one
                try
                {
                    // FeatureFillet3 with all required parameters
                    Feature filletFeature3 = (Feature)swModel.FeatureManager.FeatureFillet3(
                        195,                  // Options
                        radius / 1000.0,      // Default radius
                        0,                    // Fillet type
                        0,                    // Overflow type
                        0,                    // RadType
                        0,                    // UseAutoSelect
                        0,                    // Continuity
                        0,                    // SetBackDistance
                        0,                    // PointRadiusDistance
                        0,                    // CornerType
                        0,                    // ReverseFillet
                        0,                    // VectorReverseFillet
                        0,                    // ReverseSurfaceFillet
                        0                     // VectorReverseSurfaceFillet
                    );

                    return filletFeature3;
                }
                catch
                {
                    return null;
                }
            }
        }

        private Feature AddHoleFeature(Dictionary<string, object> parameters)
        {
            // This is a simplified hole creation
            // In real implementation, you'd need face selection and positioning

            double diameter = 10.0; // Default 10mm
            double depth = 20.0;    // Default 20mm

            if (parameters.ContainsKey("diameter"))
            {
                diameter = Convert.ToDouble(parameters["diameter"]);
            }
            if (parameters.ContainsKey("depth"))
            {
                depth = Convert.ToDouble(parameters["depth"]);
            }

            // Note: This is simplified. Real implementation needs:
            // 1. Select face
            // 2. Define position
            // 3. Create hole wizard feature

            // For now, return null as placeholder
            return null;
        }

        #endregion

        #region Save Operations

        /// <summary>
        /// Saves all modified documents in the current session
        /// </summary>
        private void SaveAllModifiedDocuments()
        {
            try
            {
                if (swModel == null) return;

                // Save the assembly
                int errors = 0;
                int warnings = 0;

                bool saveResult = swModel.Save3(
                    (int)swSaveAsOptions_e.swSaveAsOptions_Silent |
                    (int)swSaveAsOptions_e.swSaveAsOptions_SaveReferenced,
                    ref errors,
                    ref warnings
                );

                if (!saveResult)
                {
                    throw new Exception($"Failed to save assembly. Errors: {errors}, Warnings: {warnings}");
                }

                // Also save all modified components
                if (swAssembly != null)
                {
                    object[] components = (object[])swAssembly.GetComponents(false);

                    foreach (Component2 comp in components)
                    {
                        ModelDoc2 compDoc = (ModelDoc2)comp.GetModelDoc2();
                        if (compDoc != null && compDoc.GetSaveFlag())
                        {
                            compDoc.Save3(
                                (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                                ref errors,
                                ref warnings
                            );
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error saving modifications: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Saves assembly to new location with all references
        /// </summary>
        public SaveAssemblyResult SaveAssemblyAs(string newPath, bool copyReferencedFiles = true)
        {
            try
            {
                var result = new SaveAssemblyResult
                {
                    SavedFiles = new List<string>()
                };

                // Create directory if it doesn't exist
                string directory = Path.GetDirectoryName(newPath);
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                if (copyReferencedFiles)
                {
                    // Use Pack and Go for complete assembly copy
                    PackAndGo packAndGo = swModel.Extension.GetPackAndGo();

                    packAndGo.IncludeDrawings = false;
                    packAndGo.IncludeSimulationResults = false;
                    packAndGo.IncludeToolboxComponents = true;
                    packAndGo.IncludeSuppressed = true;

                    // Set destination
                    packAndGo.SetSaveToName(true, directory);

                    // Get all files that will be copied
                    object fileNames;
                    object fileStatus;
                    packAndGo.GetDocumentNames(out fileNames);
                    packAndGo.GetDocumentSaveToNames(out fileNames, out fileStatus);

                    // Perform Pack and Go
                    int[] statuses = (int[])swModel.Extension.SavePackAndGo(packAndGo);

                    string[] savedFiles = (string[])fileNames;
                    result.SavedFiles.AddRange(savedFiles);
                }
                else
                {
                    // Just save the assembly
                    int errors = 0;
                    int warnings = 0;

                    bool success = swModel.Extension.SaveAs(
                        newPath,
                        (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
                        (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                        null,
                        ref errors,
                        ref warnings
                    );

                    if (!success)
                    {
                        result.Success = false;
                        result.ErrorMessage = $"Save failed. Errors: {errors}, Warnings: {warnings}";
                        return result;
                    }

                    result.SavedFiles.Add(newPath);
                }

                // Calculate total size
                result.TotalSize = 0;
                foreach (var file in result.SavedFiles)
                {
                    if (File.Exists(file))
                    {
                        result.TotalSize += new FileInfo(file).Length;
                    }
                }

                result.Success = true;
                return result;
            }
            catch (Exception ex)
            {
                return new SaveAssemblyResult
                {
                    Success = false,
                    ErrorMessage = ex.Message
                };
            }
        }

        #endregion


        #region Helper Methods

        /// <summary>
        /// Closes all open documents
        /// </summary>
        public void CloseAllDocuments()
        {
            try
            {
                if (swApp != null)
                {
                    // Close all documents
                    swApp.CloseAllDocuments(true);

                    // Clear references
                    swModel = null;
                    swAssembly = null;
                }
            }
            catch (Exception ex)
            {
                // Log but don't throw - we want to continue even if close fails
                System.Diagnostics.Debug.WriteLine($"Error closing documents: {ex.Message}");
            }
        }


        public void CloseDocument()
        {
            if (swModel != null)
            {
                try
                {
                    swApp.CloseDoc(swModel.GetTitle());
                }
                catch
                {
                    // Ignore errors during close
                }
                finally
                {
                    swModel = null;
                    swAssembly = null;
                }
            }
        }

        private int GetComponentCount()
        {
            if (swAssembly == null) return 0;

            object[] components = (object[])swAssembly.GetComponents(false);
            return components?.Length ?? 0;
        }

        private List<string> GetFeatureList()
        {
            var features = new List<string>();

            Feature feat = (Feature)swModel.FirstFeature();
            while (feat != null)
            {
                if (!feat.GetTypeName2().StartsWith("Reference"))
                {
                    features.Add($"{feat.Name} ({feat.GetTypeName2()})");
                }
                feat = (Feature)feat.GetNextFeature();
            }

            return features;
        }

        private List<SketchInfo> GetSketchList()
        {
            var sketches = new List<SketchInfo>();

            Feature feat = (Feature)swModel.FirstFeature();
            while (feat != null)
            {
                if (feat.GetTypeName2() == "ProfileFeature")
                {
                    var sketch = (Sketch)feat.GetSpecificFeature2();
                    if (sketch != null)
                    {
                        object[] segments = (object[])sketch.GetSketchSegments();
                        int segmentCount = segments?.Length ?? 0;

                        sketches.Add(new SketchInfo
                        {
                            Name = feat.Name,
                            SegmentCount = segmentCount,
                            IsActive = sketch.Is3D()
                        });
                    }
                }
                feat = (Feature)feat.GetNextFeature();
            }

            return sketches;
        }

        private DisplayDimension FindDimension(string featureName, string dimensionName)
        {
            Feature feat = null;

            // Find feature by name
            Feature tempFeat = (Feature)swModel.FirstFeature();
            while (tempFeat != null)
            {
                if (tempFeat.Name == featureName)
                {
                    feat = tempFeat;
                    break;
                }
                tempFeat = (Feature)tempFeat.GetNextFeature();
            }

            if (feat == null) return null;

            DisplayDimension dispDim = (DisplayDimension)feat.GetFirstDisplayDimension();
            while (dispDim != null)
            {
                Dimension dim = (Dimension)dispDim.GetDimension();
                if (dim.FullName == dimensionName || dim.Name == dimensionName)
                {
                    return dispDim;
                }
                dispDim = (DisplayDimension)feat.GetNextDisplayDimension(dispDim);
            }

            return null;
        }

        private Component2 FindComponent(string componentName)
        {
            object[] components = (object[])swAssembly.GetComponents(false);

            foreach (Component2 comp in components)
            {
                if (comp.Name2 == componentName ||
                    comp.GetSelectByIDString().Contains(componentName))
                {
                    return comp;
                }
            }

            return null;
        }

        private string GetDimensionUnits(Dimension dim)
        {
            // Get document units
            int lengthUnit = swModel.GetUserPreferenceIntegerValue(
                (int)swUserPreferenceIntegerValue_e.swUnitsLinear
            );

            switch (lengthUnit)
            {
                case (int)swLengthUnit_e.swMM: return "mm";
                case (int)swLengthUnit_e.swCM: return "cm";
                case (int)swLengthUnit_e.swMETER: return "m";
                case (int)swLengthUnit_e.swINCHES: return "in";
                case (int)swLengthUnit_e.swFEET: return "ft";
                default: return "unknown";
            }
        }

        private string ExtractThumbnail(string filePath)
        {
            // Thumbnail extraction requires either:
            // 1. Document Manager API (separate license)
            // 2. Taking a screenshot of the current view
            // 3. Using Windows Shell to extract embedded thumbnails

            // For now, return null - implement this based on your specific needs
            // The UI will work fine without thumbnails

            /* Example implementation options:
            
            // Option 1: Save current view as image (requires open document)
            if (swModel != null)
            {
                string tempPath = Path.Combine(Path.GetTempPath(), "temp_thumb.bmp");
                swModel.SaveBMP(tempPath, 200, 200);
                return tempPath;
            }
            
            // Option 2: Use Document Manager API
            // Requires separate license and SwDocumentMgr reference
            
            // Option 3: Extract Windows thumbnail
            // Use Shell32 or Windows API
            */

            return null;
        }

        private string FindMaterialDatabase(string materialName)
        {
            // Common SolidWorks material database paths
            string swPath = swApp.GetExecutablePath();
            string swDir = System.IO.Path.GetDirectoryName(swPath);

            // Try standard material database locations
            string[] possiblePaths = new string[]
            {
                Path.Combine(swDir, @"lang\english\sldmaterials\SolidWorks Materials.sldmat"),
                Path.Combine(swDir, @"lang\english\sldmaterials\Custom Materials.sldmat"),
                Path.Combine(swDir, @"sldmaterials\SolidWorks Materials.sldmat")
            };

            // Check which database file exists
            foreach (var path in possiblePaths)
            {
                if (File.Exists(path))
                {
                    // In production, you'd check if the material exists in this database
                    return path;
                }
            }

            // Return empty if not found
            return "";
        }

        private void SelectEdge(string edgeName)
        {
            // This is a simplified implementation
            // In a real implementation, you would:
            // 1. Parse the edge reference (e.g., "Edge<1>")
            // 2. Find the actual edge in the model
            // 3. Select it using SelectByID2

            try
            {
                // Example: Select edge by name
                // The actual edge name format depends on your model
                swModel.Extension.SelectByID2(
                    edgeName,           // Name
                    "EDGE",            // Type
                    0, 0, 0,           // X, Y, Z coordinates (not used for edges)
                    true,              // Append to selection
                    0,                 // Mark
                    null,              // Callout
                    0                  // Selection option
                );
            }
            catch
            {
                // Log error in production
            }
        }

        private List<(string name, double value, string axis)> ExtractDimensions(ModelDoc2 model)
        {
            var dims = new List<(string, double, string)>();
            Feature feature = (Feature)model.FirstFeature();

            while (feature != null)
            {
                string type = feature.GetTypeName2();
                if (type == "Sketch" || type == "ProfileFeature")
                {
                    DisplayDimension dispDim = (DisplayDimension)feature.GetFirstDisplayDimension();
                    while (dispDim != null)
                    {
                        try
                        {
                            Dimension dim = (Dimension)dispDim.GetDimension();
                            if (dim != null)
                            {
                                string name = dim.FullName;
                                double value = dim.SystemValue;
                                int dimType = dim.GetType();

                                string axis;
                                if (dimType == 0) axis = "X";
                                else if (dimType == 1) axis = "Y";
                                else if (dimType == 2) axis = "Z";
                                else axis = "?";

                                dims.Add((name, value, axis));
                            }
                        }
                        catch { }

                        dispDim = (DisplayDimension)feature.GetNextDisplayDimension(dispDim);
                    }
                }
                feature = (Feature)feature.GetNextFeature();
            }

            return dims;
        }


        private List<Modification> ParseModificationJson(string json)
        {
            var modifications = new List<Modification>();

            try
            {
                using (var document = System.Text.Json.JsonDocument.Parse(json))
                {
                    var root = document.RootElement;

                    if (root.TryGetProperty("modifications", out var modsElement))
                    {
                        foreach (var modElement in modsElement.EnumerateArray())
                        {
                            var mod = new Modification();

                            if (modElement.TryGetProperty("type", out var typeElement))
                                mod.Type = typeElement.GetString();

                            if (mod.Type == "scale")
                            {
                                mod.Parameters = new Dictionary<string, object>();

                                if (modElement.TryGetProperty("axis", out var axisElement))
                                    mod.Parameters["axis"] = axisElement.GetString();

                                if (modElement.TryGetProperty("targetSize", out var sizeElement))
                                    mod.Parameters["targetSize"] = sizeElement.GetDouble();
                            }
                            else if (mod.Type == "dimension")
                            {
                                if (modElement.TryGetProperty("feature", out var featureElement))
                                    mod.FeatureName = featureElement.GetString();

                                if (modElement.TryGetProperty("dimension", out var dimElement))
                                    mod.DimensionName = dimElement.GetString();

                                if (modElement.TryGetProperty("newValue", out var newValElement))
                                    mod.NewValue = newValElement.GetDouble();
                            }
                            // Add other modification types as needed

                            modifications.Add(mod);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to parse JSON: {ex.Message}");
            }

            return modifications;
        }

        /// <summary>
        /// Copies assembly and all referenced files to a folder
        /// </summary>
        public CopyAssemblyResult CopyAssemblyToFolder(string assemblyPath, string destinationFolder)
        {
            try
            {
                var result = new CopyAssemblyResult
                {
                    CopiedFiles = new List<string>()
                };

                // Open assembly temporarily to get Pack and Go
                int errors = 0;
                int warnings = 0;
                ModelDoc2 tempModel = swApp.OpenDoc6(
                    assemblyPath,
                    (int)swDocumentTypes_e.swDocASSEMBLY,
                    (int)swOpenDocOptions_e.swOpenDocOptions_Silent | (int)swOpenDocOptions_e.swOpenDocOptions_ReadOnly,
                    "",
                    ref errors,
                    ref warnings
                );

                if (tempModel == null)
                {
                    result.Success = false;
                    result.ErrorMessage = "Failed to open assembly for copying";
                    return result;
                }

                try
                {
                    // Use Pack and Go to copy all files
                    PackAndGo packAndGo = tempModel.Extension.GetPackAndGo();

                    packAndGo.IncludeDrawings = false;
                    packAndGo.IncludeSimulationResults = false;
                    packAndGo.IncludeToolboxComponents = true;
                    packAndGo.IncludeSuppressed = true;

                    // Set destination
                    packAndGo.SetSaveToName(true, destinationFolder);

                    // Get all files that will be copied
                    object fileNames;
                    object fileStatus;
                    packAndGo.GetDocumentNames(out fileNames);
                    packAndGo.GetDocumentSaveToNames(out fileNames, out fileStatus);

                    // Perform Pack and Go
                    int[] statuses = (int[])tempModel.Extension.SavePackAndGo(packAndGo);

                    string[] copiedFiles = (string[])fileNames;
                    result.CopiedFiles.AddRange(copiedFiles);

                    result.Success = true;
                }
                finally
                {
                    // Close the temp model
                    swApp.CloseDoc(tempModel.GetTitle());
                }

                return result;
            }
            catch (Exception ex)
            {
                return new CopyAssemblyResult
                {
                    Success = false,
                    ErrorMessage = ex.Message
                };
            }
        }

        // Add result class at the bottom with other result classes
        public class CopyAssemblyResult
        {
            public bool Success { get; set; }
            public List<string> CopiedFiles { get; set; }
            public string ErrorMessage { get; set; }
        }


        #endregion
    }

    #region Result Classes

    public class AssemblyInfo
    {
        public string Name { get; set; }
        public string FilePath { get; set; }
        public long FileSize { get; set; }
        public int PartCount { get; set; }
        public string ThumbnailPath { get; set; }
        public List<string> Features { get; set; }
        public List<SketchInfo> Sketches { get; set; }
    }

    public class SketchInfo
    {
        public string Name { get; set; }
        public int SegmentCount { get; set; }
        public bool IsActive { get; set; }
    }

    public class ModificationResult
    {
        public string FeatureName { get; set; }
        public bool Success { get; set; }
        public string ErrorMessage { get; set; }
        public string OldValue { get; set; }
        public string NewValue { get; set; }
    }

    public class DimensionModificationResult
    {
        public bool Success { get; set; }
        public double OldValue { get; set; }
        public double NewValue { get; set; }
        public string Units { get; set; }
        public string ErrorMessage { get; set; }
    }

    public class MaterialChangeResult
    {
        public bool Success { get; set; }
        public string OldMaterial { get; set; }
        public string NewMaterial { get; set; }
        public string ErrorMessage { get; set; }
    }

    public class FeatureAddResult
    {
        public bool Success { get; set; }
        public string FeatureName { get; set; }
        public string ErrorMessage { get; set; }
    }

    public class SaveAssemblyResult
    {
        public bool Success { get; set; }
        public List<string> SavedFiles { get; set; }
        public long TotalSize { get; set; }
        public string ErrorMessage { get; set; }
    }

    public class DimensionInfo
    {
        public string ComponentName { get; set; }
        public string DimensionName { get; set; }
        public double Value { get; set; }
        public string Axis { get; set; }
    }

    public class ScalingResult
    {
        public bool Success { get; set; }
        public string ErrorMessage { get; set; }
        public double ScaleRatio { get; set; }
        public List<string> ModifiedDimensions { get; set; }
    }



    #endregion
}