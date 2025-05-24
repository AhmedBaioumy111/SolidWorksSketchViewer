using System;
using System.Collections.Generic;
using System.Text.Json;
using System.Linq;

namespace SolidWorksSketchViewer.Services
{
    /// <summary>
    /// Service for JSON processing and validation
    /// </summary>
    public class JsonProcessingService
    {
        /// <summary>
        /// Validates modification JSON structure
        /// </summary>
        public JsonValidationResult ValidateModificationJson(string jsonString)
        {
            var result = new JsonValidationResult
            {
                IsValid = true,
                ErrorMessages = new List<string>(),
                ParsedModifications = null
            };

            try
            {
                // Check if JSON is empty
                if (string.IsNullOrWhiteSpace(jsonString))
                {
                    result.IsValid = false;
                    result.ErrorMessages.Add("JSON is empty");
                    return result;
                }

                // Try to parse JSON
                using (JsonDocument document = JsonDocument.Parse(jsonString))
                {
                    var root = document.RootElement;

                    // Check for modifications array
                    if (!root.TryGetProperty("modifications", out JsonElement modifications))
                    {
                        result.IsValid = false;
                        result.ErrorMessages.Add("Missing 'modifications' array");
                        return result;
                    }

                    if (modifications.ValueKind != JsonValueKind.Array)
                    {
                        result.IsValid = false;
                        result.ErrorMessages.Add("'modifications' must be an array");
                        return result;
                    }

                    // Parse and validate each modification
                    var modificationSet = new ModificationSet
                    {
                        Modifications = new List<Modification>()
                    };

                    foreach (var modElement in modifications.EnumerateArray())
                    {
                        var mod = ParseModification(modElement, result.ErrorMessages);
                        if (mod != null)
                        {
                            modificationSet.Modifications.Add(mod);
                        }
                    }

                    result.ParsedModifications = modificationSet;

                    // Check if we have at least one valid modification
                    if (modificationSet.Modifications.Count == 0)
                    {
                        result.IsValid = false;
                        result.ErrorMessages.Add("No valid modifications found");
                    }
                }
            }
            catch (JsonException ex)
            {
                result.IsValid = false;
                result.ErrorMessages.Add($"JSON parsing error: {ex.Message}");
            }
            catch (Exception ex)
            {
                result.IsValid = false;
                result.ErrorMessages.Add($"Validation error: {ex.Message}");
            }

            return result;
        }

        /// <summary>
        /// Converts modifications to JSON
        /// </summary>
        public string ConvertToJson(List<IModification> modifications)
        {
            try
            {
                var modificationObjects = new List<object>();

                foreach (var mod in modifications)
                {
                    if (mod is DimensionModification dimMod)
                    {
                        modificationObjects.Add(new
                        {
                            type = "dimension",
                            feature = dimMod.FeatureName,
                            dimension = dimMod.DimensionName,
                            currentValue = dimMod.CurrentValue,
                            newValue = dimMod.NewValue,
                            units = dimMod.Units
                        });
                    }
                    else if (mod is MaterialModification matMod)
                    {
                        modificationObjects.Add(new
                        {
                            type = "material",
                            component = matMod.ComponentName,
                            currentMaterial = matMod.CurrentMaterial,
                            newMaterial = matMod.NewMaterial
                        });
                    }
                    else if (mod is FeatureModification featMod)
                    {
                        modificationObjects.Add(new
                        {
                            type = "feature",
                            operation = featMod.Operation,
                            featureType = featMod.FeatureType,
                            edges = featMod.Edges,
                            parameters = featMod.Parameters
                        });
                    }
                }

                var jsonObject = new { modifications = modificationObjects };

                return JsonSerializer.Serialize(jsonObject, new JsonSerializerOptions
                {
                    WriteIndented = true,
                    PropertyNamingPolicy = JsonNamingPolicy.CamelCase
                });
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to convert to JSON: {ex.Message}", ex);
            }
        }

        private Modification ParseModification(JsonElement element, List<string> errors)
        {
            try
            {
                // Get modification type
                if (!element.TryGetProperty("type", out JsonElement typeElement))
                {
                    errors.Add("Modification missing 'type' property");
                    return null;
                }

                string type = typeElement.GetString()?.ToLower();

                switch (type)
                {
                    case "dimension":
                        return ParseDimensionModification(element, errors);

                    case "material":
                        return ParseMaterialModification(element, errors);

                    case "feature":
                        return ParseFeatureModification(element, errors);

                    default:
                        errors.Add($"Unknown modification type: {type}");
                        return null;
                }
            }
            catch (Exception ex)
            {
                errors.Add($"Error parsing modification: {ex.Message}");
                return null;
            }
        }

        private Modification ParseDimensionModification(JsonElement element, List<string> errors)
        {
            var mod = new Modification
            {
                Type = "dimension",
                Parameters = new Dictionary<string, object>()
            };

            // Required fields
            if (element.TryGetProperty("feature", out JsonElement feature))
                mod.FeatureName = feature.GetString();
            else
                errors.Add("Dimension modification missing 'feature'");

            if (element.TryGetProperty("dimension", out JsonElement dimension))
                mod.DimensionName = dimension.GetString();

            if (element.TryGetProperty("newValue", out JsonElement newValue))
                mod.NewValue = newValue.GetDouble();
            else
                errors.Add("Dimension modification missing 'newValue'");

            // Optional fields
            if (element.TryGetProperty("currentValue", out JsonElement currentValue))
                mod.CurrentValue = currentValue.GetDouble();

            if (element.TryGetProperty("units", out JsonElement units))
                mod.Units = units.GetString();

            return mod;
        }

        private Modification ParseMaterialModification(JsonElement element, List<string> errors)
        {
            var mod = new Modification
            {
                Type = "material",
                Parameters = new Dictionary<string, object>()
            };

            // Required fields
            if (element.TryGetProperty("component", out JsonElement component))
                mod.FeatureName = component.GetString();
            else if (element.TryGetProperty("feature", out JsonElement feature))
                mod.FeatureName = feature.GetString();
            else
                errors.Add("Material modification missing 'component' or 'feature'");

            if (element.TryGetProperty("newMaterial", out JsonElement newMaterial))
                mod.NewValue = newMaterial.GetString();
            else if (element.TryGetProperty("newValue", out JsonElement newValue))
                mod.NewValue = newValue.GetString();
            else
                errors.Add("Material modification missing 'newMaterial'");

            // Optional fields
            if (element.TryGetProperty("currentMaterial", out JsonElement currentMaterial))
                mod.CurrentValue = currentMaterial.GetString();

            return mod;
        }

        private Modification ParseFeatureModification(JsonElement element, List<string> errors)
        {
            var mod = new Modification
            {
                Type = "feature",
                Parameters = new Dictionary<string, object>()
            };

            // Required fields
            if (element.TryGetProperty("operation", out JsonElement operation))
                mod.Parameters["operation"] = operation.GetString();
            else
                errors.Add("Feature modification missing 'operation'");

            if (element.TryGetProperty("edges", out JsonElement edges) && edges.ValueKind == JsonValueKind.Array)
            {
                var edgeList = new List<string>();
                foreach (var edge in edges.EnumerateArray())
                {
                    edgeList.Add(edge.GetString());
                }
                mod.Parameters["edges"] = edgeList;
            }

            // Parse additional parameters
            if (element.TryGetProperty("value", out JsonElement value))
                mod.Parameters["value"] = value.GetDouble();

            if (element.TryGetProperty("parameters", out JsonElement parameters))
            {
                foreach (var prop in parameters.EnumerateObject())
                {
                    mod.Parameters[prop.Name] = ParseJsonValue(prop.Value);
                }
            }

            return mod;
        }

        private object ParseJsonValue(JsonElement element)
        {
            switch (element.ValueKind)
            {
                case JsonValueKind.String:
                    return element.GetString();
                case JsonValueKind.Number:
                    return element.GetDouble();
                case JsonValueKind.True:
                case JsonValueKind.False:
                    return element.GetBoolean();
                case JsonValueKind.Array:
                    var list = new List<object>();
                    foreach (var item in element.EnumerateArray())
                    {
                        list.Add(ParseJsonValue(item));
                    }
                    return list;
                case JsonValueKind.Object:
                    var dict = new Dictionary<string, object>();
                    foreach (var prop in element.EnumerateObject())
                    {
                        dict[prop.Name] = ParseJsonValue(prop.Value);
                    }
                    return dict;
                default:
                    return null;
            }
        }
    }

    /// <summary>
    /// Result of JSON validation
    /// </summary>
    public class JsonValidationResult
    {
        public bool IsValid { get; set; }
        public List<string> ErrorMessages { get; set; }
        public ModificationSet ParsedModifications { get; set; }
    }

    /// <summary>
    /// Set of modifications
    /// </summary>
    public class ModificationSet
    {
        public List<Modification> Modifications { get; set; }
    }

    /// <summary>
    /// Base modification class
    /// Note: This should be moved to a shared Models folder to avoid duplication
    /// </summary>
    public class Modification
    {
        public string Type { get; set; }
        public string FeatureName { get; set; }
        public string DimensionName { get; set; }
        public object CurrentValue { get; set; }
        public object NewValue { get; set; }
        public string Units { get; set; }
        public Dictionary<string, object> Parameters { get; set; }
    }

    /// <summary>
    /// Interface for modifications
    /// </summary>
    public interface IModification
    {
        string Type { get; }
        string GetDescription();
    }

    /// <summary>
    /// Dimension modification
    /// </summary>
    public class DimensionModification : IModification
    {
        public string Type => "dimension";
        public string FeatureName { get; set; }
        public string DimensionName { get; set; }
        public double CurrentValue { get; set; }
        public double NewValue { get; set; }
        public string Units { get; set; }

        public string GetDescription()
        {
            return $"Change {DimensionName} from {CurrentValue}{Units} to {NewValue}{Units}";
        }
    }

    /// <summary>
    /// Material modification
    /// </summary>
    public class MaterialModification : IModification
    {
        public string Type => "material";
        public string ComponentName { get; set; }
        public string CurrentMaterial { get; set; }
        public string NewMaterial { get; set; }

        public string GetDescription()
        {
            return $"Change material of {ComponentName} from {CurrentMaterial} to {NewMaterial}";
        }
    }

    /// <summary>
    /// Feature modification
    /// </summary>
    public class FeatureModification : IModification
    {
        public string Type => "feature";
        public string Operation { get; set; }
        public string FeatureType { get; set; }
        public List<string> Edges { get; set; }
        public Dictionary<string, object> Parameters { get; set; }

        public string GetDescription()
        {
            return $"{Operation} {FeatureType} on {Edges?.Count ?? 0} edges";
        }
    }
}