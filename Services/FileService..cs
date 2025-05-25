using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using SolidWorksSketchViewer.Models;

namespace SolidWorksSketchViewer.Services
{
    /// <summary>
    /// Service for all file operations
    /// </summary>
    public class FileService
    {
        /// <summary>
        /// Scans directory for SolidWorks assembly files
        /// </summary>
        public List<FileItemModel> GetAssemblyFiles(string directoryPath)
        {
            var files = new List<FileItemModel>();

            try
            {
                if (!Directory.Exists(directoryPath))
                {
                    return files;
                }

                // Get all assembly files
                var assemblyFiles = Directory.GetFiles(directoryPath, "*.sldasm", SearchOption.AllDirectories)
                    .Union(Directory.GetFiles(directoryPath, "*.SLDASM", SearchOption.AllDirectories));

                foreach (var file in assemblyFiles)
                {
                    var fileInfo = new FileInfo(file);
                    files.Add(new FileItemModel
                    {
                        FileName = fileInfo.Name,
                        FilePath = fileInfo.FullName,
                        FileSize = FormatFileSize(fileInfo.Length),
                        LastModified = fileInfo.LastWriteTime,
                        ValidationStatusColor = ValidateFile(file) ? "Green" : "Orange"
                    });
                }
            }
            catch (Exception ex)
            {
                // Log error
                Console.WriteLine($"Error scanning directory: {ex.Message}");
            }

            return files;
        }

        /// <summary>
        /// Scans directory for BOM files (Excel, CSV)
        /// </summary>
        public List<FileItemModel> GetBOMFiles(string directoryPath)
        {
            var files = new List<FileItemModel>();

            try
            {
                if (!Directory.Exists(directoryPath))
                    return files;

                // Get Excel and CSV files
                var bomFiles = Directory.GetFiles(directoryPath, "*.xlsx", SearchOption.AllDirectories)
                    .Union(Directory.GetFiles(directoryPath, "*.xls", SearchOption.AllDirectories))
                    .Union(Directory.GetFiles(directoryPath, "*.csv", SearchOption.AllDirectories));

                foreach (var file in bomFiles)
                {
                    var fileInfo = new FileInfo(file);
                    files.Add(new FileItemModel
                    {
                        FileName = fileInfo.Name,
                        FilePath = fileInfo.FullName,
                        FileSize = FormatFileSize(fileInfo.Length),
                        LastModified = fileInfo.LastWriteTime,
                        ValidationStatusColor = "Green"
                    });
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error scanning for BOM files: {ex.Message}");
            }

            return files;
        }

        /// <summary>
        /// Scans directory for Requirements files (txt, docx, json)
        /// </summary>
        public List<FileItemModel> GetRequirementsFiles(string directoryPath)
        {
            var files = new List<FileItemModel>();

            try
            {
                if (!Directory.Exists(directoryPath))
                    return files;

                // Get text, Word, and JSON files
                var reqFiles = Directory.GetFiles(directoryPath, "*.txt", SearchOption.AllDirectories)
                    .Union(Directory.GetFiles(directoryPath, "*.docx", SearchOption.AllDirectories))
                    .Union(Directory.GetFiles(directoryPath, "*.json", SearchOption.AllDirectories));

                foreach (var file in reqFiles)
                {
                    var fileInfo = new FileInfo(file);
                    files.Add(new FileItemModel
                    {
                        FileName = fileInfo.Name,
                        FilePath = fileInfo.FullName,
                        FileSize = FormatFileSize(fileInfo.Length),
                        LastModified = fileInfo.LastWriteTime,
                        ValidationStatusColor = "Green"
                    });
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error scanning for requirements files: {ex.Message}");
            }

            return files;
        }
        /// <summary>
        /// Validates SolidWorks assembly file
        /// </summary>
        public FileValidationResult ValidateAssemblyFile(string filePath)
        {
            var result = new FileValidationResult
            {
                IsValid = true,
                StatusColor = "Green",
                ErrorMessages = new List<string>()
            };

            try
            {
                if (!File.Exists(filePath))
                {
                    result.IsValid = false;
                    result.StatusColor = "Red";
                    result.ErrorMessages.Add("File does not exist");
                    return result;
                }

                var fileInfo = new FileInfo(filePath);

                // Check file size
                if (fileInfo.Length == 0)
                {
                    result.IsValid = false;
                    result.StatusColor = "Red";
                    result.ErrorMessages.Add("File is empty");
                }
                else if (fileInfo.Length > 500 * 1024 * 1024) // > 500MB
                {
                    result.StatusColor = "Orange";
                    result.ErrorMessages.Add("Large file - may take time to process");
                }

                // Check if file is locked
                if (IsFileLocked(filePath))
                {
                    result.StatusColor = "Orange";
                    result.ErrorMessages.Add("File is currently in use");
                }

                // Check file extension
                string extension = fileInfo.Extension.ToLower();
                if (extension != ".sldasm")
                {
                    result.StatusColor = "Orange";
                    result.ErrorMessages.Add("Non-standard file extension");
                }
            }
            catch (Exception ex)
            {
                result.IsValid = false;
                result.StatusColor = "Red";
                result.ErrorMessages.Add($"Validation error: {ex.Message}");
            }

            return result;
        }

        /// <summary>
        /// Reads BOM file and returns structured data
        /// </summary>
        public List<BOMPreviewItem> ReadBOMFile(string filePath)
        {
            var items = new List<BOMPreviewItem>();

            try
            {
                string extension = Path.GetExtension(filePath).ToLower();

                if (extension == ".csv")
                {
                    // Read CSV file
                    var lines = File.ReadAllLines(filePath);
                    bool isHeader = true;

                    foreach (var line in lines)
                    {
                        if (isHeader)
                        {
                            isHeader = false;
                            continue;
                        }

                        var parts = line.Split(',');
                        if (parts.Length >= 4)
                        {
                            items.Add(new BOMPreviewItem
                            {
                                PartName = parts[0].Trim('"'),
                                Quantity = int.TryParse(parts[1], out int qty) ? qty : 1,
                                Material = parts[2].Trim('"'),
                                Description = parts[3].Trim('"')
                            });
                        }
                    }
                }
                else if (extension == ".xlsx" || extension == ".xls")
                {
                    // For Excel files, you would use a library like EPPlus or ClosedXML
                    // For now, return mock data
                    items.Add(new BOMPreviewItem
                    {
                        PartName = "Motor Mount",
                        Quantity = 1,
                        Material = "Aluminum 6061",
                        Description = "Main motor mounting bracket"
                    });
                    items.Add(new BOMPreviewItem
                    {
                        PartName = "Bearing Housing",
                        Quantity = 2,
                        Material = "Steel 1045",
                        Description = "Bearing support structure"
                    });
                }
            }
            catch (Exception ex)
            {
                // Log error and return empty list
                Console.WriteLine($"Error reading BOM file: {ex.Message}");
            }

            return items;
        }

        /// <summary>
        /// Reads requirements text file
        /// </summary>
        public string ReadRequirementsFile(string filePath)
        {
            try
            {
                if (!File.Exists(filePath))
                {
                    return "File not found";
                }

                string extension = Path.GetExtension(filePath).ToLower();

                switch (extension)
                {
                    case ".txt":
                        return File.ReadAllText(filePath);

                    case ".json":
                        // Pretty print JSON
                        string json = File.ReadAllText(filePath);
                        return FormatJson(json);

                    case ".docx":
                        // For Word files, you would use a library like DocX
                        // For now, return placeholder
                        return "Word document support coming soon. Please use .txt or .json files.";

                    default:
                        return $"Unsupported file type: {extension}";
                }
            }
            catch (Exception ex)
            {
                return $"Error reading file: {ex.Message}";
            }
        }

        #region Helper Methods

        private string FormatFileSize(long bytes)
        {
            string[] sizes = { "B", "KB", "MB", "GB" };
            double len = bytes;
            int order = 0;

            while (len >= 1024 && order < sizes.Length - 1)
            {
                order++;
                len = len / 1024;
            }

            return $"{len:0.#} {sizes[order]}";
        }

        private bool ValidateFile(string filePath)
        {
            try
            {
                var fileInfo = new FileInfo(filePath);
                return fileInfo.Exists && fileInfo.Length > 0 && !IsFileLocked(filePath);
            }
            catch
            {
                return false;
            }
        }

        private bool IsFileLocked(string filePath)
        {
            try
            {
                using (FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.None))
                {
                    stream.Close();
                }
                return false;
            }
            catch (IOException)
            {
                return true;
            }
        }

        private string FormatJson(string json)
        {
            try
            {
                // Simple JSON formatting
                var parsedJson = System.Text.Json.JsonSerializer.Deserialize<object>(json);
                return System.Text.Json.JsonSerializer.Serialize(parsedJson, new System.Text.Json.JsonSerializerOptions
                {
                    WriteIndented = true
                });
            }
            catch
            {
                return json;
            }
        }

        #endregion
    }

    /// <summary>
    /// File validation result
    /// </summary>
    public class FileValidationResult
    {
        public bool IsValid { get; set; }
        public string StatusColor { get; set; }
        public List<string> ErrorMessages { get; set; }
    }
}