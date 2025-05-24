using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using SolidWorksSketchViewer.Models;

namespace SolidWorksSketchViewer.Services
{
    /// <summary>
    /// Service for LLM/AI integration
    /// </summary>
    public class LLMService
    {
        private Random _random = new Random();

        /// <summary>
        /// Processes requirements text using AI
        /// </summary>
        public async Task<LLMAnalysisResult> AnalyzeRequirements(
            string requirementsText,
            AssemblyContext assemblyContext = null,
            Action<ProcessingStep> progressCallback = null)
        {
            var result = new LLMAnalysisResult
            {
                ExtractedRequirements = new List<ExtractedRequirement>(),
                FeatureMappings = new List<FeatureMapping>(),
                Conflicts = new List<ConflictItem>(),
                ModificationJSON = ""
            };

            try
            {
                // Simulate processing steps
                var steps = new[]
                {
                    "Parsing requirements text...",
                    "Extracting dimensions and constraints...",
                    "Mapping to SolidWorks features...",
                    "Validating modifications...",
                    "Generating modification JSON..."
                };

                foreach (var step in steps)
                {
                    progressCallback?.Invoke(new ProcessingStep
                    {
                        Status = "⏳",
                        Message = step
                    });

                    await Task.Delay(500); // Simulate processing time

                    // In real implementation, this is where you'd call your LLM API
                    // For example:
                    // var response = await CallOpenAI(requirementsText, step);
                }

                // Parse requirements (mock implementation)
                result.ExtractedRequirements = ExtractRequirementsFromText(requirementsText);

                // Map to features (mock implementation)
                result.FeatureMappings = MapRequirementsToFeatures(result.ExtractedRequirements, assemblyContext);

                // Check for conflicts
                result.Conflicts = CheckForConflicts(result.FeatureMappings);

                // Generate JSON
                result.ModificationJSON = GenerateModificationJson(result.FeatureMappings);

                // Final success callback
                progressCallback?.Invoke(new ProcessingStep
                {
                    Status = "✓",
                    Message = "Analysis complete"
                });
            }
            catch (Exception ex)
            {
                progressCallback?.Invoke(new ProcessingStep
                {
                    Status = "❌",
                    Message = $"Error: {ex.Message}"
                });
                throw;
            }

            return result;
        }

        private List<ExtractedRequirement> ExtractRequirementsFromText(string text)
        {
            var requirements = new List<ExtractedRequirement>();

            // Mock implementation - parse common patterns
            var lines = text.Split('\n');

            foreach (var line in lines)
            {
                var lowerLine = line.ToLower();

                // Look for dimension changes
                if (lowerLine.Contains("diameter") || lowerLine.Contains("radius") ||
                    lowerLine.Contains("length") || lowerLine.Contains("width") ||
                    lowerLine.Contains("height"))
                {
                    // Extract numbers using regex (simplified)
                    var numbers = ExtractNumbers(line);
                    if (numbers.Count >= 2)
                    {
                        requirements.Add(new ExtractedRequirement
                        {
                            Text = line.Trim(),
                            Type = "Dimension",
                            Confidence = 85 + _random.Next(15)
                        });
                    }
                }
                // Look for material changes
                else if (lowerLine.Contains("material") || lowerLine.Contains("aluminum") ||
                         lowerLine.Contains("steel") || lowerLine.Contains("plastic"))
                {
                    requirements.Add(new ExtractedRequirement
                    {
                        Text = line.Trim(),
                        Type = "Material",
                        Confidence = 80 + _random.Next(20)
                    });
                }
                // Look for feature additions
                else if (lowerLine.Contains("chamfer") || lowerLine.Contains("fillet") ||
                         lowerLine.Contains("hole") || lowerLine.Contains("groove"))
                {
                    requirements.Add(new ExtractedRequirement
                    {
                        Text = line.Trim(),
                        Type = "Feature",
                        Confidence = 70 + _random.Next(25)
                    });
                }
                // Look for constraints
                else if (lowerLine.Contains("minimum") || lowerLine.Contains("maximum") ||
                         lowerLine.Contains("clearance") || lowerLine.Contains("thickness"))
                {
                    requirements.Add(new ExtractedRequirement
                    {
                        Text = line.Trim(),
                        Type = "Constraint",
                        Confidence = 60 + _random.Next(30)
                    });
                }
            }

            return requirements;
        }

        private List<FeatureMapping> MapRequirementsToFeatures(
            List<ExtractedRequirement> requirements,
            AssemblyContext context)
        {
            var mappings = new List<FeatureMapping>();

            foreach (var req in requirements)
            {
                var mapping = new FeatureMapping
                {
                    Requirement = req.Text,
                    Status = "Valid"
                };

                switch (req.Type)
                {
                    case "Dimension":
                        // Map to sketch dimensions
                        mapping.TargetFeature = $"Sketch{_random.Next(1, 5)} - D{_random.Next(1, 3)}@Sketch{_random.Next(1, 5)}";
                        mapping.CurrentValue = $"{10 + _random.Next(20)}mm";
                        mapping.NewValue = ExtractFirstNumber(req.Text).ToString() + "mm";
                        break;

                    case "Material":
                        mapping.TargetFeature = "Part Properties";
                        mapping.CurrentValue = "Steel 1045";
                        mapping.NewValue = ExtractMaterial(req.Text);
                        break;

                    case "Feature":
                        mapping.TargetFeature = $"Edge<{_random.Next(1, 10)}>";
                        mapping.CurrentValue = "None";
                        mapping.NewValue = ExtractFeatureType(req.Text);
                        break;

                    case "Constraint":
                        mapping.TargetFeature = "Assembly Constraints";
                        mapping.CurrentValue = "Not defined";
                        mapping.NewValue = ExtractConstraintValue(req.Text);
                        mapping.Status = "Warning";
                        break;
                }

                mappings.Add(mapping);
            }

            return mappings;
        }

        private List<ConflictItem> CheckForConflicts(List<FeatureMapping> mappings)
        {
            var conflicts = new List<ConflictItem>();

            // Check for potential conflicts
            var edgeMappings = mappings.Where(m => m.TargetFeature.Contains("Edge")).ToList();
            if (edgeMappings.Count > 3)
            {
                conflicts.Add(new ConflictItem
                {
                    Title = "Multiple Edge Modifications",
                    Description = $"{edgeMappings.Count} edges will be modified. This may cause conflicts with existing features.",
                    Resolution = "Consider applying changes incrementally or verify edge references."
                });
            }

            // Check for dimension conflicts
            var dimMappings = mappings.Where(m => m.TargetFeature.Contains("Sketch")).ToList();
            foreach (var dim in dimMappings)
            {
                if (dim.NewValue.Contains("mm"))
                {
                    double value = ExtractFirstNumber(dim.NewValue);
                    if (value > 100)
                    {
                        conflicts.Add(new ConflictItem
                        {
                            Title = "Large Dimension Change",
                            Description = $"Dimension {dim.TargetFeature} will be changed to {value}mm. This is a significant change.",
                            Resolution = "Verify this dimension change won't cause assembly interference."
                        });
                    }
                }
            }

            return conflicts;
        }

        private string GenerateModificationJson(List<FeatureMapping> mappings)
        {
            var modifications = new List<object>();

            foreach (var mapping in mappings)
            {
                if (mapping.TargetFeature.Contains("Sketch"))
                {
                    modifications.Add(new
                    {
                        type = "dimension",
                        feature = mapping.TargetFeature.Split('-')[0].Trim(),
                        dimension = mapping.TargetFeature.Split('-')[1].Trim(),
                        currentValue = ExtractFirstNumber(mapping.CurrentValue),
                        newValue = ExtractFirstNumber(mapping.NewValue),
                        units = "mm"
                    });
                }
                else if (mapping.TargetFeature == "Part Properties")
                {
                    modifications.Add(new
                    {
                        type = "material",
                        property = "Material",
                        currentValue = mapping.CurrentValue,
                        newValue = mapping.NewValue
                    });
                }
                else if (mapping.TargetFeature.Contains("Edge"))
                {
                    modifications.Add(new
                    {
                        type = "feature",
                        operation = "add_" + mapping.NewValue.ToLower(),
                        edges = new[] { mapping.TargetFeature },
                        value = 5.0,
                        units = "mm"
                    });
                }
            }

            var json = System.Text.Json.JsonSerializer.Serialize(new { modifications },
                new System.Text.Json.JsonSerializerOptions { WriteIndented = true });

            return json;
        }

        #region Helper Methods

        private List<double> ExtractNumbers(string text)
        {
            var numbers = new List<double>();
            var words = text.Split(' ');

            foreach (var word in words)
            {
                var cleaned = word.Trim('m', 'M', ',', '.', '(', ')');
                if (double.TryParse(cleaned, out double number))
                {
                    numbers.Add(number);
                }
            }

            return numbers;
        }

        private double ExtractFirstNumber(string text)
        {
            var numbers = ExtractNumbers(text);
            return numbers.Count > 0 ? numbers[0] : 10.0;
        }

        private string ExtractMaterial(string text)
        {
            var lowerText = text.ToLower();

            if (lowerText.Contains("aluminum") || lowerText.Contains("aluminium"))
            {
                if (lowerText.Contains("6061"))
                    return "Aluminum 6061";
                else
                    return "Aluminum Alloy";
            }
            else if (lowerText.Contains("steel"))
            {
                if (lowerText.Contains("stainless"))
                    return "Stainless Steel 316";
                else
                    return "Steel 1045";
            }
            else if (lowerText.Contains("plastic") || lowerText.Contains("abs"))
            {
                return "ABS Plastic";
            }

            return "Aluminum 6061"; // Default
        }

        private string ExtractFeatureType(string text)
        {
            var lowerText = text.ToLower();

            if (lowerText.Contains("chamfer"))
                return "Chamfer";
            else if (lowerText.Contains("fillet"))
                return "Fillet";
            else if (lowerText.Contains("hole"))
                return "Hole";
            else if (lowerText.Contains("groove"))
                return "Groove";

            return "Feature";
        }

        private string ExtractConstraintValue(string text)
        {
            var numbers = ExtractNumbers(text);
            if (numbers.Count > 0)
                return $"{numbers[0]}mm";
            return "3mm"; // Default
        }

        #endregion
    }

    /// <summary>
    /// Context information about the assembly
    /// </summary>
    public class AssemblyContext
    {
        public string AssemblyName { get; set; }
        public List<string> AvailableFeatures { get; set; }
        public List<string> CurrentDimensions { get; set; }
    }

    /// <summary>
    /// Result of LLM analysis
    /// </summary>
    public class LLMAnalysisResult
    {
        public List<ExtractedRequirement> ExtractedRequirements { get; set; }
        public List<FeatureMapping> FeatureMappings { get; set; }
        public List<ConflictItem> Conflicts { get; set; }
        public string ModificationJSON { get; set; }
    }
}