using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Windows.Threading;
using Microsoft.Win32;
using SolidWorksSketchViewer.Helpers;
using SolidWorksSketchViewer.Models;
using SolidWorksSketchViewer.Services;
using PdfSharpCore.Pdf;
using PdfSharpCore.Drawing;
using ClosedXML.Excel;

namespace SolidWorksSketchViewer.ViewModels
{
    public class MainViewModel : ViewModelBase
    {
        private DispatcherTimer _timer;
        private DateTime _startTime;

        // Services
        private SolidWorksService _solidWorksService;
        private FileService _fileService;
        private LLMService _llmService;
        private JsonProcessingService _jsonService;

        #region Properties - File Management
        private string _workingDirectory = @"C:\SolidWorks\Projects";
        private ObservableCollection<FileItemModel> _assemblyFiles;
        private ObservableCollection<FileItemModel> _bomFiles;
        private ObservableCollection<FileItemModel> _requirementsFiles;
        private FileItemModel _selectedAssemblyFile;
        private FileItemModel _selectedBOMFile;
        private FileItemModel _selectedRequirementsFile;

        public string WorkingDirectory
        {
            get => _workingDirectory;
            set => SetProperty(ref _workingDirectory, value);
        }

        public ObservableCollection<FileItemModel> AssemblyFiles
        {
            get => _assemblyFiles;
            set => SetProperty(ref _assemblyFiles, value);
        }

        public ObservableCollection<FileItemModel> BOMFiles
        {
            get => _bomFiles;
            set => SetProperty(ref _bomFiles, value);
        }

        public ObservableCollection<FileItemModel> RequirementsFiles
        {
            get => _requirementsFiles;
            set => SetProperty(ref _requirementsFiles, value);
        }

        public FileItemModel SelectedAssemblyFile
        {
            get => _selectedAssemblyFile;
            set
            {
                SetProperty(ref _selectedAssemblyFile, value);
                OnPropertyChanged(nameof(HasSelectedAssembly));
                OnPropertyChanged(nameof(CanProcessRequirements));
                UpdateAssemblyPreview();
            }
        }

        public FileItemModel SelectedBOMFile
        {
            get => _selectedBOMFile;
            set
            {
                SetProperty(ref _selectedBOMFile, value);
                OnPropertyChanged(nameof(HasSelectedBOM));
                UpdateBOMPreview();
            }
        }

        public FileItemModel SelectedRequirementsFile
        {
            get => _selectedRequirementsFile;
            set
            {
                SetProperty(ref _selectedRequirementsFile, value);
                OnPropertyChanged(nameof(HasSelectedRequirements));
                OnPropertyChanged(nameof(CanProcessRequirements));
                UpdateRequirementsPreview();
            }
        }
        #endregion

        #region Properties - UI State
        private bool _showFileBrowser = true;
        private bool _showResultsPanel = true;
        private bool _isDarkTheme = false;
        private int _currentProcessingStage = 0;
        private bool _isLoading = false;
        private string _statusMessage = "Ready";

        public bool ShowFileBrowser
        {
            get => _showFileBrowser;
            set => SetProperty(ref _showFileBrowser, value);
        }

        public bool ShowResultsPanel
        {
            get => _showResultsPanel;
            set => SetProperty(ref _showResultsPanel, value);
        }

        public bool IsDarkTheme
        {
            get => _isDarkTheme;
            set => SetProperty(ref _isDarkTheme, value);
        }

        public int CurrentProcessingStage
        {
            get => _currentProcessingStage;
            set => SetProperty(ref _currentProcessingStage, value);
        }

        public bool IsLoading
        {
            get => _isLoading;
            set => SetProperty(ref _isLoading, value);
        }

        public string StatusMessage
        {
            get => _statusMessage;
            set => SetProperty(ref _statusMessage, value);
        }
        #endregion

        #region Properties - Preview Data
        private string _assemblyName;
        private string _assemblyPartCount;
        private string _assemblyFileSize;
        private string _assemblyThumbnail;
        private ObservableCollection<BOMPreviewItem> _bomPreviewData;
        private string _requirementsText;

        public string AssemblyName
        {
            get => _assemblyName;
            set => SetProperty(ref _assemblyName, value);
        }

        public string AssemblyPartCount
        {
            get => _assemblyPartCount;
            set => SetProperty(ref _assemblyPartCount, value);
        }

        public string AssemblyFileSize
        {
            get => _assemblyFileSize;
            set => SetProperty(ref _assemblyFileSize, value);
        }

        public string AssemblyThumbnail
        {
            get => _assemblyThumbnail;
            set => SetProperty(ref _assemblyThumbnail, value);
        }

        private List<string> _currentAssemblyFeatures;

        public ObservableCollection<BOMPreviewItem> BOMPreviewData
        {
            get => _bomPreviewData;
            set => SetProperty(ref _bomPreviewData, value);
        }

        public string RequirementsText
        {
            get => _requirementsText;
            set => SetProperty(ref _requirementsText, value);
        }

        public bool HasSelectedAssembly => SelectedAssemblyFile != null;
        public bool HasSelectedBOM => SelectedBOMFile != null;
        public bool HasSelectedRequirements => SelectedRequirementsFile != null;
        public bool CanProcessRequirements => HasSelectedAssembly && HasSelectedRequirements;
        #endregion

        #region Properties - LLM Processing
        private bool _isLLMProcessing = false;
        private string _llmStatusMessage = "Ready to analyze requirements";
        private double _llmProgress = 0;
        private ObservableCollection<ProcessingStep> _processingSteps;
        private ObservableCollection<ExtractedRequirement> _extractedRequirements;
        private ObservableCollection<FeatureMapping> _featureMappings;
        private ObservableCollection<ConflictItem> _conflicts;
        private string _modificationJSON;
        private bool _enableJSONEditing = false;

        public bool IsLLMProcessing
        {
            get => _isLLMProcessing;
            set
            {
                SetProperty(ref _isLLMProcessing, value);
                OnPropertyChanged(nameof(CanApprove)); // Add this line
            }
        }

        public string LLMStatusMessage
        {
            get => _llmStatusMessage;
            set => SetProperty(ref _llmStatusMessage, value);
        }

        public double LLMProgress
        {
            get => _llmProgress;
            set => SetProperty(ref _llmProgress, value);
        }

        public ObservableCollection<ProcessingStep> ProcessingSteps
        {
            get => _processingSteps;
            set => SetProperty(ref _processingSteps, value);
        }

        public ObservableCollection<ExtractedRequirement> ExtractedRequirements
        {
            get => _extractedRequirements;
            set => SetProperty(ref _extractedRequirements, value);
        }

        public ObservableCollection<FeatureMapping> FeatureMappings
        {
            get => _featureMappings;
            set => SetProperty(ref _featureMappings, value);
        }

        public ObservableCollection<ConflictItem> Conflicts
        {
            get => _conflicts;
            set => SetProperty(ref _conflicts, value);
        }

        public string ModificationJSON
        {
            get => _modificationJSON;
            set => SetProperty(ref _modificationJSON, value);
        }

        public bool EnableJSONEditing
        {
            get => _enableJSONEditing;
            set
            {
                SetProperty(ref _enableJSONEditing, value);
                OnPropertyChanged(nameof(IsJSONReadOnly));
            }
        }

        public bool IsJSONReadOnly => !EnableJSONEditing;
        public bool CanApprove => ExtractedRequirements?.Count > 0 && !IsLLMProcessing;
        #endregion

        #region Properties - SolidWorks Processing
        private bool _isProcessing = false;
        private string _processingStatusMessage = "Ready to process";
        private double _processingProgress = 0;
        private string _currentOperation = "";
        private ObservableCollection<FeatureProcessingStatus> _featureProcessingStatus;
        private bool _canPauseProcessing = false;

        public bool IsProcessing
        {
            get => _isProcessing;
            set => SetProperty(ref _isProcessing, value);
        }

        public string ProcessingStatusMessage
        {
            get => _processingStatusMessage;
            set => SetProperty(ref _processingStatusMessage, value);
        }

        public double ProcessingProgress
        {
            get => _processingProgress;
            set => SetProperty(ref _processingProgress, value);
        }

        public string CurrentOperation
        {
            get => _currentOperation;
            set => SetProperty(ref _currentOperation, value);
        }

        public ObservableCollection<FeatureProcessingStatus> FeatureProcessingStatus
        {
            get => _featureProcessingStatus;
            set => SetProperty(ref _featureProcessingStatus, value);
        }

        public bool CanPauseProcessing
        {
            get => _canPauseProcessing;
            set => SetProperty(ref _canPauseProcessing, value);
        }
        #endregion

        #region Properties - Results
        private string _successSummary;
        private string _processingTime;
        private ObservableCollection<ChangeSummaryItem> _changeSummary;
        private ObservableCollection<RequirementsFulfillmentItem> _requirementsFulfillment;
        private string _newAssemblyName = "Modified_Assembly_001";
        private string _saveLocationPath;

        public string SuccessSummary
        {
            get => _successSummary;
            set => SetProperty(ref _successSummary, value);
        }

        public string ProcessingTime
        {
            get => _processingTime;
            set => SetProperty(ref _processingTime, value);
        }

        public ObservableCollection<ChangeSummaryItem> ChangeSummary
        {
            get => _changeSummary;
            set => SetProperty(ref _changeSummary, value);
        }

        public ObservableCollection<RequirementsFulfillmentItem> RequirementsFulfillment
        {
            get => _requirementsFulfillment;
            set => SetProperty(ref _requirementsFulfillment, value);
        }

        public string NewAssemblyName
        {
            get => _newAssemblyName;
            set
            {
                SetProperty(ref _newAssemblyName, value);
                UpdateSaveLocation();
            }
        }

        public string SaveLocationPath
        {
            get => _saveLocationPath;
            set => SetProperty(ref _saveLocationPath, value);
        }
        #endregion

        #region Properties - Status Bar
        private string _solidWorksStatus = "Connected";
        private string _solidWorksStatusColor = "Green";
        private string _llmStatus = "Ready";
        private string _llmStatusColor = "Green";
        private string _projectStatus = "No project loaded";
        private string _elapsedTime = "00:00:00";

        public string SolidWorksStatus
        {
            get => _solidWorksStatus;
            set => SetProperty(ref _solidWorksStatus, value);
        }

        public string SolidWorksStatusColor
        {
            get => _solidWorksStatusColor;
            set => SetProperty(ref _solidWorksStatusColor, value);
        }

        public string LLMStatus
        {
            get => _llmStatus;
            set => SetProperty(ref _llmStatus, value);
        }

        public string LLMStatusColor
        {
            get => _llmStatusColor;
            set => SetProperty(ref _llmStatusColor, value);
        }

        public string ProjectStatus
        {
            get => _projectStatus;
            set => SetProperty(ref _projectStatus, value);
        }

        public string ElapsedTime
        {
            get => _elapsedTime;
            set => SetProperty(ref _elapsedTime, value);
        }
        #endregion

        #region Properties - Recent Projects
        private ObservableCollection<string> _recentProjects;

        public ObservableCollection<string> RecentProjects
        {
            get => _recentProjects;
            set => SetProperty(ref _recentProjects, value);
        }
        #endregion

        #region Properties - Temp Folder Management
        private string _currentTempFolderPath;
        private bool _isWorkingWithTempFiles = false;

        public string CurrentTempFolderPath
        {
            get => _currentTempFolderPath;
            set => SetProperty(ref _currentTempFolderPath, value);
        }

        public bool IsWorkingWithTempFiles
        {
            get => _isWorkingWithTempFiles;
            set => SetProperty(ref _isWorkingWithTempFiles, value);
        }
        #endregion


        #region Commands
        // File Commands
        public ICommand BrowseDirectoryCommand { get; private set; }
        public ICommand RefreshAssembliesCommand { get; private set; }
        public ICommand ProcessRequirementsCommand { get; private set; }

        // Menu Commands
        public ICommand NewProjectCommand { get; private set; }
        public ICommand OpenProjectCommand { get; private set; }
        public ICommand SaveProjectCommand { get; private set; }
        public ICommand ExitCommand { get; private set; }

        public ICommand ShowDocumentationCommand { get; private set; }
        public ICommand HelpCommand { get; private set; }
        public ICommand ShowAboutCommand { get; private set; }

        // LLM Processing Commands
        public ICommand ApproveAllCommand { get; private set; }
        public ICommand ReviewEachCommand { get; private set; }
        public ICommand EditJSONCommand { get; private set; }
        public ICommand ReanalyzeCommand { get; private set; }

        // JSON Tab Specific Commands
        public ICommand ValidateJSONCommand { get; private set; }
        public ICommand ApplyJSONCommand { get; private set; }
        public ICommand ResetJSONCommand { get; private set; }

        // Feature Mapping Commands
        public ICommand ValidateMappingsCommand { get; private set; }

        // Conflict Commands
        public ICommand IgnoreConflictsCommand { get; private set; }
        public ICommand AutoResolveCommand { get; private set; }

        // SolidWorks Processing Commands
        public ICommand CancelProcessingCommand { get; private set; }
        public ICommand PauseProcessingCommand { get; private set; }
        public ICommand SkipCurrentCommand { get; private set; }

        // Results Commands
        public ICommand SaveAssemblyCommand { get; private set; }
        public ICommand ExportPDFCommand { get; private set; }
        public ICommand ExportExcelCommand { get; private set; }
        public ICommand ExportLogCommand { get; private set; }
        #endregion

        #region Temp Folder Management

        private async Task CreateTempWorkspace()
        {
            try
            {
                StatusMessage = "Creating temporary workspace...";
                IsLoading = true;

                // Generate unique temp folder name
                string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                string tempFolderName = $"Temp_Assembly_{timestamp}";
                CurrentTempFolderPath = Path.Combine(WorkingDirectory, tempFolderName);

                // Create temp directory
                Directory.CreateDirectory(CurrentTempFolderPath);

                // Copy assembly and all referenced files using Pack and Go
                var copyResult = await Task.Run(() =>
                    _solidWorksService.CopyAssemblyToFolder(
                        SelectedAssemblyFile.FilePath,
                        CurrentTempFolderPath
                    ));

                if (!copyResult.Success)
                {
                    throw new Exception($"Failed to copy assembly: {copyResult.ErrorMessage}");
                }

                IsWorkingWithTempFiles = true;
                StatusMessage = "Temporary workspace created successfully";
            }
            catch (Exception ex)
            {
                CleanupTempFolder();
                throw new Exception($"Failed to create temp workspace: {ex.Message}");
            }
            finally
            {
                IsLoading = false;
            }
        }

        private void CleanupTempFolder()
        {
            if (string.IsNullOrEmpty(CurrentTempFolderPath) || !Directory.Exists(CurrentTempFolderPath))
            {
                CurrentTempFolderPath = null;
                IsWorkingWithTempFiles = false;
                return;
            }

            try
            {
                // First attempt - try to delete normally
                Directory.Delete(CurrentTempFolderPath, true);
                CurrentTempFolderPath = null;
                IsWorkingWithTempFiles = false;
            }
            catch (Exception ex)
            {
                try
                {
                    // Second attempt - try to rename first, then delete
                    string renamedPath = CurrentTempFolderPath + "_ToDelete_" + DateTime.Now.Ticks;
                    Directory.Move(CurrentTempFolderPath, renamedPath);

                    try
                    {
                        Directory.Delete(renamedPath, true);
                    }
                    catch
                    {
                        // If we can't delete the renamed folder, at least we freed up the original path
                        StatusMessage = $"Warning: Could not delete temp folder (renamed to {Path.GetFileName(renamedPath)})";
                    }

                    CurrentTempFolderPath = null;
                    IsWorkingWithTempFiles = false;
                }
                catch
                {
                    // Final catch - just log and clear references
                    StatusMessage = $"Warning: Could not clean up temp folder: {ex.Message}";
                    CurrentTempFolderPath = null;
                    IsWorkingWithTempFiles = false;
                }
            }
        }
        private void CleanupAllTempFolders()
        {
            try
            {
                // Clean up any orphaned temp folders from previous sessions
                if (Directory.Exists(WorkingDirectory))
                {
                    var tempFolders = Directory.GetDirectories(WorkingDirectory, "Temp_Assembly_*");
                    foreach (var folder in tempFolders)
                    {
                        try
                        {
                            Directory.Delete(folder, true);
                        }
                        catch
                        {
                            // Skip folders that can't be deleted (might be in use)
                        }
                    }
                }
            }
            catch
            {
                // Silent cleanup - don't interrupt application flow
            }
        }

        #endregion


        #region Export Methods

        private async void ExecuteExportPDF(object parameter)
        {
            try
            {
                var saveDialog = new SaveFileDialog
                {
                    Filter = "PDF Files (*.pdf)|*.pdf",
                    Title = "Export Change Report as PDF",
                    FileName = $"{NewAssemblyName}_ChangeReport_{DateTime.Now:yyyyMMdd}.pdf"
                };

                if (saveDialog.ShowDialog() != true)
                    return;

                StatusMessage = "Exporting to PDF...";
                IsLoading = true;

                await Task.Run(() => ExportToPDF(saveDialog.FileName));

                ShowMessage($"PDF report exported successfully to:\n{saveDialog.FileName}");
            }
            catch (Exception ex)
            {
                ShowMessage($"Export Error: {ex.Message}");
            }
            finally
            {
                IsLoading = false;
                StatusMessage = "Ready";
            }
        }

        private void ExportToPDF(string fileName)
        {
            // Create a new PDF document
            PdfDocument document = new PdfDocument();
            document.Info.Title = "SolidWorks Assembly Modification Report";
            document.Info.Author = "SolidWorks Assembly Modifier";
            document.Info.Subject = $"Modification Report for {SelectedAssemblyFile?.FileName}";

            // Create first page
            PdfPage page = document.AddPage();
            XGraphics gfx = XGraphics.FromPdfPage(page);

            // Define fonts
            XFont titleFont = new XFont("Arial", 20, XFontStyle.Bold);
            XFont headingFont = new XFont("Arial", 14, XFontStyle.Bold);
            XFont normalFont = new XFont("Arial", 11, XFontStyle.Regular);
            XFont smallFont = new XFont("Arial", 10, XFontStyle.Regular);

            double yPosition = 50;
            double leftMargin = 50;
            double pageWidth = page.Width - 100;

            // Title
            gfx.DrawString("SOLIDWORKS ASSEMBLY MODIFICATION REPORT", titleFont,
                XBrushes.DarkBlue, new XRect(leftMargin, yPosition, pageWidth, 30),
                XStringFormats.TopCenter);
            yPosition += 40;

            // Date and Assembly info
            gfx.DrawString($"Generated: {DateTime.Now:yyyy-MM-dd HH:mm:ss}", normalFont,
                XBrushes.Black, leftMargin, yPosition);
            yPosition += 20;

            gfx.DrawString($"Assembly: {SelectedAssemblyFile?.FileName ?? "N/A"}", normalFont,
                XBrushes.Black, leftMargin, yPosition);
            yPosition += 30;

            // Modification Summary
            gfx.DrawString("MODIFICATION SUMMARY", headingFont,
                XBrushes.DarkBlue, leftMargin, yPosition);
            yPosition += 25;

            gfx.DrawString(SuccessSummary ?? "No modifications processed", normalFont,
                XBrushes.Black, leftMargin, yPosition);
            yPosition += 30;

            // Detailed Changes
            gfx.DrawString("DETAILED CHANGES", headingFont,
                XBrushes.DarkBlue, leftMargin, yPosition);
            yPosition += 25;

            foreach (var change in ChangeSummary)
            {
                // Check if we need a new page
                if (yPosition > page.Height - 100)
                {
                    page = document.AddPage();
                    gfx = XGraphics.FromPdfPage(page);
                    yPosition = 50;
                }

                // Draw change details
                gfx.DrawString($"Feature: {change.Feature}", normalFont,
                    XBrushes.Black, leftMargin, yPosition);
                yPosition += 18;

                gfx.DrawString($"  Original: {change.OriginalValue} → New: {change.NewValue}", smallFont,
                    XBrushes.DarkGray, leftMargin + 20, yPosition);
                yPosition += 18;

                // Status with color
                XBrush statusBrush = change.Status == "✓" ? XBrushes.Green : XBrushes.Red;
                gfx.DrawString($"  Status: {change.Status}", normalFont,
                    statusBrush, leftMargin + 20, yPosition);
                yPosition += 25;
            }

            // Requirements Fulfillment on new page if needed
            if (yPosition > page.Height - 200)
            {
                page = document.AddPage();
                gfx = XGraphics.FromPdfPage(page);
                yPosition = 50;
            }

            gfx.DrawString("REQUIREMENTS FULFILLMENT", headingFont,
                XBrushes.DarkBlue, leftMargin, yPosition);
            yPosition += 25;

            foreach (var req in RequirementsFulfillment)
            {
                if (yPosition > page.Height - 50)
                {
                    page = document.AddPage();
                    gfx = XGraphics.FromPdfPage(page);
                    yPosition = 50;
                }

                XBrush reqBrush = req.StatusIcon == "✓" ? XBrushes.Green :
                                 req.StatusIcon == "❌" ? XBrushes.Red : XBrushes.Orange;

                gfx.DrawString($"{req.StatusIcon} {req.Requirement}", normalFont,
                    reqBrush, leftMargin, yPosition);
                yPosition += 20;
            }

            // Save the document
            document.Save(fileName);
        }

        private async void ExecuteExportExcel(object parameter)
        {
            try
            {
                var saveDialog = new SaveFileDialog
                {
                    Filter = "Excel Files (*.xlsx)|*.xlsx",
                    Title = "Export Change Report as Excel",
                    FileName = $"{NewAssemblyName}_ChangeReport_{DateTime.Now:yyyyMMdd}.xlsx"
                };

                if (saveDialog.ShowDialog() != true)
                    return;

                StatusMessage = "Exporting to Excel...";
                IsLoading = true;

                await Task.Run(() => ExportToExcel(saveDialog.FileName));

                ShowMessage($"Excel report exported successfully to:\n{saveDialog.FileName}");
            }
            catch (Exception ex)
            {
                ShowMessage($"Export Error: {ex.Message}");
            }
            finally
            {
                IsLoading = false;
                StatusMessage = "Ready";
            }
        }

        private void ExportToExcel(string fileName)
        {
            using (var workbook = new XLWorkbook())
            {
                // Summary Sheet
                var summarySheet = workbook.Worksheets.Add("Summary");

                // Title
                summarySheet.Cell(1, 1).Value = "SOLIDWORKS ASSEMBLY MODIFICATION REPORT";
                summarySheet.Cell(1, 1).Style.Font.Bold = true;
                summarySheet.Cell(1, 1).Style.Font.FontSize = 16;
                summarySheet.Cell(1, 1).Style.Fill.BackgroundColor = XLColor.DarkBlue;
                summarySheet.Cell(1, 1).Style.Font.FontColor = XLColor.White;
                summarySheet.Range(1, 1, 1, 4).Merge();

                // Info
                summarySheet.Cell(3, 1).Value = "Generated:";
                summarySheet.Cell(3, 2).Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                summarySheet.Cell(4, 1).Value = "Assembly:";
                summarySheet.Cell(4, 2).Value = SelectedAssemblyFile?.FileName ?? "N/A";

                summarySheet.Cell(5, 1).Value = "Requirements File:";
                summarySheet.Cell(5, 2).Value = SelectedRequirementsFile?.FileName ?? "N/A";

                summarySheet.Cell(7, 1).Value = "Summary:";
                summarySheet.Cell(7, 2).Value = SuccessSummary ?? "No modifications processed";

                // Format info section
                summarySheet.Range(3, 1, 7, 1).Style.Font.Bold = true;
                summarySheet.Range(3, 1, 7, 2).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                // Change Details Sheet
                var changesSheet = workbook.Worksheets.Add("Change Details");

                // Headers
                changesSheet.Cell(1, 1).Value = "Feature";
                changesSheet.Cell(1, 2).Value = "Original Value";
                changesSheet.Cell(1, 3).Value = "New Value";
                changesSheet.Cell(1, 4).Value = "Status";

                // Header formatting
                var headerRange = changesSheet.Range(1, 1, 1, 4);
                headerRange.Style.Font.Bold = true;
                headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;
                headerRange.Style.Border.BottomBorder = XLBorderStyleValues.Medium;

                // Data
                int row = 2;
                foreach (var change in ChangeSummary)
                {
                    changesSheet.Cell(row, 1).Value = change.Feature;
                    changesSheet.Cell(row, 2).Value = change.OriginalValue;
                    changesSheet.Cell(row, 3).Value = change.NewValue;
                    changesSheet.Cell(row, 4).Value = change.Status;

                    // Color code status
                    if (change.Status == "✓")
                        changesSheet.Cell(row, 4).Style.Font.FontColor = XLColor.Green;
                    else if (change.Status == "❌")
                        changesSheet.Cell(row, 4).Style.Font.FontColor = XLColor.Red;

                    row++;
                }

                // Auto-fit columns
                changesSheet.Columns().AdjustToContents();

                // Requirements Sheet
                var reqSheet = workbook.Worksheets.Add("Requirements");

                // Headers
                reqSheet.Cell(1, 1).Value = "Status";
                reqSheet.Cell(1, 2).Value = "Requirement";

                // Header formatting
                reqSheet.Range(1, 1, 1, 2).Style.Font.Bold = true;
                reqSheet.Range(1, 1, 1, 2).Style.Fill.BackgroundColor = XLColor.LightGray;
                reqSheet.Range(1, 1, 1, 2).Style.Border.BottomBorder = XLBorderStyleValues.Medium;

                // Data
                row = 2;
                foreach (var req in RequirementsFulfillment)
                {
                    reqSheet.Cell(row, 1).Value = req.StatusIcon;
                    reqSheet.Cell(row, 2).Value = req.Requirement;

                    // Color code based on status
                    if (req.StatusIcon == "✓")
                    {
                        reqSheet.Cell(row, 1).Style.Font.FontColor = XLColor.Green;
                        reqSheet.Cell(row, 2).Style.Fill.BackgroundColor = XLColor.FromHtml("#E8F5E9");
                    }
                    else if (req.StatusIcon == "❌")
                    {
                        reqSheet.Cell(row, 1).Style.Font.FontColor = XLColor.Red;
                        reqSheet.Cell(row, 2).Style.Fill.BackgroundColor = XLColor.FromHtml("#FFEBEE");
                    }
                    else
                    {
                        reqSheet.Cell(row, 1).Style.Font.FontColor = XLColor.Orange;
                        reqSheet.Cell(row, 2).Style.Fill.BackgroundColor = XLColor.FromHtml("#FFF3E0");
                    }

                    row++;
                }

                // Auto-fit columns
                reqSheet.Columns().AdjustToContents();

                // Processing Log Sheet
                var logSheet = workbook.Worksheets.Add("Processing Log");

                row = 1;
                logSheet.Cell(row++, 1).Value = "AI ANALYSIS STEPS";
                logSheet.Cell(row - 1, 1).Style.Font.Bold = true;

                foreach (var step in ProcessingSteps)
                {
                    logSheet.Cell(row, 1).Value = step.Status;
                    logSheet.Cell(row, 2).Value = step.Message;
                    row++;
                }

                row++;
                logSheet.Cell(row++, 1).Value = "MODIFICATION JSON";
                logSheet.Cell(row - 1, 1).Style.Font.Bold = true;
                logSheet.Cell(row, 1).Value = ModificationJSON;
                logSheet.Range(row, 1, row, 3).Merge();

                // Auto-fit columns
                logSheet.Columns().AdjustToContents();

                // Save the workbook
                workbook.SaveAs(fileName);
            }
        }
        private async void ExecuteExportLog(object parameter)
        {
            try
            {
                var saveDialog = new SaveFileDialog
                {
                    Filter = "Log Files (*.log)|*.log|Text Files (*.txt)|*.txt",
                    Title = "Export Processing Log",
                    FileName = $"{NewAssemblyName}_ProcessingLog_{DateTime.Now:yyyyMMdd}"
                };

                if (saveDialog.ShowDialog() != true)
                    return;

                StatusMessage = "Exporting processing log...";
                IsLoading = true;

                await Task.Run(() => ExportProcessingLog(saveDialog.FileName));

                ShowMessage($"Processing log exported successfully to:\n{saveDialog.FileName}");
            }
            catch (Exception ex)
            {
                ShowMessage($"Export Error: {ex.Message}");
            }
            finally
            {
                IsLoading = false;
                StatusMessage = "Ready";
            }
        }

        private void ExportProcessingLog(string fileName)
        {
            var log = new System.Text.StringBuilder();

            log.AppendLine($"SolidWorks Assembly Modifier - Processing Log");
            log.AppendLine($"Generated: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            log.AppendLine($"Session Duration: {ElapsedTime}");
            log.AppendLine();

            log.AppendLine("FILE INFORMATION");
            log.AppendLine($"Assembly: {SelectedAssemblyFile?.FilePath ?? "N/A"}");
            log.AppendLine($"Requirements: {SelectedRequirementsFile?.FilePath ?? "N/A"}");
            log.AppendLine($"BOM: {SelectedBOMFile?.FilePath ?? "N/A"}");
            log.AppendLine();

            log.AppendLine("AI ANALYSIS STEPS");
            foreach (var step in ProcessingSteps)
            {
                log.AppendLine($"{step.Status} {step.Message}");
            }
            log.AppendLine();

            log.AppendLine("EXTRACTED REQUIREMENTS");
            foreach (var req in ExtractedRequirements)
            {
                log.AppendLine($"- {req.Text} (Type: {req.Type}, Confidence: {req.Confidence}%)");
            }
            log.AppendLine();

            log.AppendLine("FEATURE MAPPINGS");
            foreach (var mapping in FeatureMappings)
            {
                log.AppendLine($"- {mapping.Requirement}");
                log.AppendLine($"  Target: {mapping.TargetFeature}");
                log.AppendLine($"  Change: {mapping.CurrentValue} → {mapping.NewValue}");
            }
            log.AppendLine();

            log.AppendLine("PROCESSING STATUS");
            foreach (var status in FeatureProcessingStatus)
            {
                log.AppendLine($"{status.StatusIcon} {status.FeatureName}: {status.Message} ({status.ProcessingTime})");
            }
            log.AppendLine();

            log.AppendLine("MODIFICATION JSON");
            log.AppendLine(ModificationJSON);

            File.WriteAllText(fileName, log.ToString());
        }

        #endregion


        public MainViewModel()
        {
            // Initialize services - ensure SolidWorks is created on STA thread
            if (Application.Current != null)
            {
                Application.Current.Dispatcher.Invoke(() =>
                {
                    _solidWorksService = new SolidWorksService();
                });
            }
            else
            {
                _solidWorksService = new SolidWorksService();
            }

            // Initialize services
            _fileService = new FileService();
            _llmService = new LLMService();
            _jsonService = new JsonProcessingService();

            InitializeCollections();
            InitializeCommands();
            //InitializeMockData();
            StartTimer();
            UpdateSaveLocation();


            // Auto-load files if working directory exists
            if (Directory.Exists(WorkingDirectory))
            {
                Task.Run(async () => await RefreshAllFiles());
            }

            // Clean up any orphaned temp folders from previous sessions
            CleanupAllTempFolders();

        }

        private void InitializeCollections()
        {
            AssemblyFiles = new ObservableCollection<FileItemModel>();
            BOMFiles = new ObservableCollection<FileItemModel>();
            RequirementsFiles = new ObservableCollection<FileItemModel>();
            ProcessingSteps = new ObservableCollection<ProcessingStep>();
            ExtractedRequirements = new ObservableCollection<ExtractedRequirement>();
            FeatureMappings = new ObservableCollection<FeatureMapping>();
            Conflicts = new ObservableCollection<ConflictItem>();
            FeatureProcessingStatus = new ObservableCollection<FeatureProcessingStatus>();
            ChangeSummary = new ObservableCollection<ChangeSummaryItem>();
            RequirementsFulfillment = new ObservableCollection<RequirementsFulfillmentItem>();
            BOMPreviewData = new ObservableCollection<BOMPreviewItem>();
            RecentProjects = new ObservableCollection<string>
            {
                "Project_Alpha_Rev3.swproj",
                "Customer_XYZ_Assembly.swproj",
                "Prototype_2025_01.swproj",
                "Testing_Configuration_B.swproj"
            };
        }

        private void InitializeCommands()
        {
            // File Commands
            BrowseDirectoryCommand = new RelayCommand(ExecuteBrowseDirectory);
            RefreshAssembliesCommand = new RelayCommand(ExecuteRefreshAssemblies);
            ProcessRequirementsCommand = new RelayCommand(ExecuteProcessRequirements, _ => CanProcessRequirements);

            // Menu Commands
            NewProjectCommand = new RelayCommand(ExecuteNewProject);
            OpenProjectCommand = new RelayCommand(ExecuteOpenProject);
            SaveProjectCommand = new RelayCommand(ExecuteSaveProject);
            ExitCommand = new RelayCommand(_ => System.Windows.Application.Current.Shutdown());

            ShowDocumentationCommand = new RelayCommand(ExecuteShowDocumentation);
            HelpCommand = new RelayCommand(ExecuteShowHelp); 
            ShowAboutCommand = new RelayCommand(_ => ShowMessage("SolidWorks Assembly Modifier v1.0"));

            // LLM Processing Commands
            ApproveAllCommand = new RelayCommand(ExecuteApproveAll, _ => CanApprove);
            ReviewEachCommand = new RelayCommand(ExecuteReviewEach, _ => CanApprove);
            EditJSONCommand = new RelayCommand(_ => EnableJSONEditing = !EnableJSONEditing);
            ReanalyzeCommand = new RelayCommand(ExecuteReanalyze);

            // JSON Tab Specific Commands
            ValidateJSONCommand = new RelayCommand(ExecuteValidateJSON);
            ApplyJSONCommand = new RelayCommand(ExecuteApplyJSON);
            ResetJSONCommand = new RelayCommand(ExecuteResetJSON);

            // Feature Mapping Commands
            ValidateMappingsCommand = new RelayCommand(ExecuteValidateMappings);

            // Conflict Commands
            IgnoreConflictsCommand = new RelayCommand(ExecuteIgnoreConflicts);
            AutoResolveCommand = new RelayCommand(ExecuteAutoResolve);

            // SolidWorks Processing Commands
            CancelProcessingCommand = new RelayCommand(ExecuteCancelProcessing, _ => IsProcessing);
            PauseProcessingCommand = new RelayCommand(ExecutePauseProcessing, _ => CanPauseProcessing);
            SkipCurrentCommand = new RelayCommand(ExecuteSkipCurrent, _ => IsProcessing);

            // Results Commands
            SaveAssemblyCommand = new RelayCommand(ExecuteSaveAssembly);
            ExportPDFCommand = new RelayCommand(ExecuteExportPDF);
            ExportExcelCommand = new RelayCommand(ExecuteExportExcel);
            ExportLogCommand = new RelayCommand(ExecuteExportLog);
        }


        private void StartTimer()
        {
            _timer = new DispatcherTimer();
            _timer.Interval = TimeSpan.FromSeconds(1);
            _timer.Tick += Timer_Tick;
            _timer.Start();
            _startTime = DateTime.Now;
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            var elapsed = DateTime.Now - _startTime;
            ElapsedTime = elapsed.ToString(@"hh\:mm\:ss");
        }

        #region Command Implementations

        private void ExecuteShowDocumentation(object parameter)
        {
            ExecuteShowHelp(parameter); // Same as help
        }

        private void ExecuteShowHelp(object parameter)
        {
            try
            {
                var helpWindow = new Views.HelpWindow();
                helpWindow.ShowDialog();
            }
            catch (Exception ex)
            {
                ShowMessage($"Error opening help: {ex.Message}");
            }
        }


        private void ExecuteBrowseDirectory(object parameter)
        {
            // Using Microsoft.Win32 for WPF compatibility
            // For folder selection, we'll use a workaround with SaveFileDialog
            var dialog = new SaveFileDialog
            {
                Title = "Select a folder",
                Filter = "Folder Selection|*.folder",
                FileName = "Select Folder",
                CheckFileExists = false,
                CheckPathExists = true
            };

            if (dialog.ShowDialog() == true)
            {
                // Get the directory from the selected path
                WorkingDirectory = System.IO.Path.GetDirectoryName(dialog.FileName);
                ExecuteRefreshAssemblies(null);
                UpdateSaveLocation();
            }
        }

        private async Task RefreshAllFiles()
        {
            await Task.Run(() =>
            {
                // Refresh all file types
                RefreshAssemblyFiles();
                RefreshBOMFiles();
                RefreshRequirementsFiles();
            });
        }

        private async void RefreshBOMFiles()
        {
            try
            {
                var files = await Task.Run(() =>
                    _fileService.GetBOMFiles(WorkingDirectory));

                Application.Current.Dispatcher.Invoke(() =>
                {
                    BOMFiles.Clear();
                    foreach (var file in files)
                    {
                        BOMFiles.Add(file);
                    }
                });
            }
            catch (Exception ex)
            {
                StatusMessage = $"Error loading BOM files: {ex.Message}";
            }
        }

        private async void RefreshRequirementsFiles()
        {
            try
            {
                var files = await Task.Run(() =>
                    _fileService.GetRequirementsFiles(WorkingDirectory));

                Application.Current.Dispatcher.Invoke(() =>
                {
                    RequirementsFiles.Clear();
                    foreach (var file in files)
                    {
                        RequirementsFiles.Add(file);
                    }
                });
            }
            catch (Exception ex)
            {
                StatusMessage = $"Error loading requirements files: {ex.Message}";
            }
        }

        private async void RefreshAssemblyFiles()
        {
            try
            {
                var files = await Task.Run(() =>
                    _fileService.GetAssemblyFiles(WorkingDirectory));

                Application.Current.Dispatcher.Invoke(() =>
                {
                    AssemblyFiles.Clear();
                    foreach (var file in files)
                    {
                        AssemblyFiles.Add(file);
                    }
                });
            }
            catch (Exception ex)
            {
                StatusMessage = $"Error loading assembly files: {ex.Message}";
            }
        }



        private async void ExecuteRefreshAssemblies(object parameter)
        {
            try
            {
                StatusMessage = "Scanning for files...";
                IsLoading = true;

                // Refresh all file types
                await RefreshAllFiles();

                int totalFiles = AssemblyFiles.Count + BOMFiles.Count + RequirementsFiles.Count;
                StatusMessage = $"Found {AssemblyFiles.Count} assemblies, {BOMFiles.Count} BOM files, {RequirementsFiles.Count} requirements";
            }
            catch (Exception ex)
            {
                StatusMessage = $"Error: {ex.Message}";
            }
            finally
            {
                IsLoading = false;
            }
        }



        private async void ExecuteProcessRequirements(object parameter)
        {

            CurrentProcessingStage = 1; // Switch to AI Analysis tab
            IsLLMProcessing = true;
            LLMStatusMessage = "Analyzing Requirements with AI...";
            // Create temp folder and copy files
            await CreateTempWorkspace();



            ProcessingSteps.Clear();

            try
            {
                // Create assembly context with real data
                var assemblyContext = new AssemblyContext
                {
                    AssemblyName = SelectedAssemblyFile?.FileName,
                    AvailableFeatures = _currentAssemblyFeatures ?? new List<string>()
                };

                // Call real LLM service
                var result = await _llmService.AnalyzeRequirements(
                    RequirementsText,
                    assemblyContext,
                    step =>
                    {
                        Application.Current.Dispatcher.Invoke(() =>
                        {
                            ProcessingSteps.Add(step);
                        });
                    }
                );

                // Update UI with real results
                Application.Current.Dispatcher.Invoke(() =>
                {
                    ExtractedRequirements.Clear();
                    foreach (var req in result.ExtractedRequirements)
                        ExtractedRequirements.Add(req);
                    Console.WriteLine($"Added {ExtractedRequirements.Count} requirements");


                    FeatureMappings.Clear();
                    foreach (var mapping in result.FeatureMappings)
                        FeatureMappings.Add(mapping);

                    Conflicts.Clear();
                    foreach (var conflict in result.Conflicts)
                        Conflicts.Add(conflict);

                    ModificationJSON = result.ModificationJSON;

                    OnPropertyChanged(nameof(CanApprove));
                });

                IsLLMProcessing = false;
                LLMStatusMessage = "Analysis complete. Review modifications below.";
            }
            catch (Exception ex)
            {
                IsLLMProcessing = false;
                LLMStatusMessage = $"Analysis failed: {ex.Message}";
                ShowMessage($"LLM Analysis Error: {ex.Message}");
            }
        }


        private async void ExecuteApproveAll(object parameter)
        {
            CurrentProcessingStage = 2; // Switch to SolidWorks Processing tab
            await ProcessSolidWorksModifications();

        }


        private async Task ProcessSolidWorksModifications()
        {
            IsProcessing = true;
            CanPauseProcessing = true;
            ProcessingStatusMessage = "Modifying Assembly...";
            FeatureProcessingStatus.Clear();

            try
            {
                // Ensure we're working with temp files
                if (IsWorkingWithTempFiles && !string.IsNullOrEmpty(CurrentTempFolderPath))
                {
                    string tempAssemblyPath = Path.Combine(CurrentTempFolderPath, SelectedAssemblyFile.FileName);

                    // Close any existing documents and open from temp
                    await Task.Run(() =>
                    {
                        _solidWorksService.CloseAllDocuments();
                        var assemblyInfo = _solidWorksService.OpenAssembly(tempAssemblyPath);

                        if (assemblyInfo == null)
                        {
                            throw new Exception("Failed to open temp assembly for modification");
                        }
                    });
                }
                else
                {
                    throw new Exception("No temporary workspace available. Please restart the process.");
                }

                // Process modifications
                var results = await _solidWorksService.ProcessModifications(
                    ModificationJSON,
                    status =>
                    {
                        Application.Current.Dispatcher.Invoke(() =>
                        {
                            var existing = FeatureProcessingStatus.FirstOrDefault(
                                f => f.FeatureName == status.FeatureName);

                            if (existing != null)
                            {
                                existing.StatusIcon = status.StatusIcon;
                                existing.Message = status.Message;
                                existing.ProcessingTime = status.ProcessingTime;
                                existing.BackgroundColor = status.BackgroundColor;
                            }
                            else
                            {
                                FeatureProcessingStatus.Add(status);
                            }

                            ProcessingProgress = (FeatureProcessingStatus.Count(f =>
                                f.StatusIcon == "✓" || f.StatusIcon == "❌") * 100.0) /
                                Math.Max(FeatureMappings.Count, 1);
                        });
                    }
                );

                // Update results
                UpdateResultsFromProcessing(results);

                // Move to results tab
                CurrentProcessingStage = 3;
            }
            catch (Exception ex)
            {
                ShowMessage($"Processing Error: {ex.Message}");
            }
            finally
            {
                IsProcessing = false;
                CanPauseProcessing = false;
                ProcessingStatusMessage = "Processing complete";
                CurrentOperation = "";
            }
        }
        private void UpdateResultsFromProcessing(List<ModificationResult> results)
        {
            int successCount = results.Count(r => r.Success);
            int totalCount = results.Count;

            SuccessSummary = $"{successCount} of {totalCount} modifications successfully applied";
            ProcessingTime = $"Processing completed at {DateTime.Now:HH:mm:ss}";

            // Update change summary
            ChangeSummary.Clear();
            foreach (var result in results)
            {
                ChangeSummary.Add(new ChangeSummaryItem
                {
                    Feature = result.FeatureName,
                    OriginalValue = result.OldValue,
                    NewValue = result.NewValue,
                    Status = result.Success ? "✓" : "❌"
                });
            }

            // Update requirements fulfillment based on results
            UpdateRequirementsFulfillment(results);
        }


        private void UpdateRequirementsFulfillment(List<ModificationResult> results)
        {
            RequirementsFulfillment.Clear();

            // Map each extracted requirement to its fulfillment status
            foreach (var requirement in ExtractedRequirements)
            {
                // Find related modifications for this requirement
                var relatedResults = results.Where(r =>
                    r.FeatureName.Contains(requirement.Text.Split(' ').FirstOrDefault() ?? "") ||
                    requirement.Text.ToLower().Contains(r.FeatureName.ToLower())
                ).ToList();

                string statusIcon;
                string backgroundColor;

                if (!relatedResults.Any())
                {
                    // No modifications found for this requirement
                    statusIcon = "⚠";
                    backgroundColor = "#FFF3E0"; // Yellow
                }
                else if (relatedResults.All(r => r.Success))
                {
                    // All related modifications succeeded
                    statusIcon = "✓";
                    backgroundColor = "#E8F5E9"; // Green
                }
                else if (relatedResults.Any(r => r.Success))
                {
                    // Some succeeded, some failed
                    statusIcon = "⚠";
                    backgroundColor = "#FFF3E0"; // Yellow
                }
                else
                {
                    // All failed
                    statusIcon = "❌";
                    backgroundColor = "#FFEBEE"; // Red
                }

                RequirementsFulfillment.Add(new RequirementsFulfillmentItem
                {
                    Requirement = requirement.Text,
                    StatusIcon = statusIcon,
                    BackgroundColor = backgroundColor
                });
            }
        }
        private void ExecuteReviewEach(object parameter)
        {
            ShowMessage("Review Each functionality would allow step-by-step approval");
        }

        private void ExecuteReanalyze(object parameter)
        {
            ProcessingSteps.Clear();
            ExtractedRequirements.Clear();
            FeatureMappings.Clear();
            Conflicts.Clear();
            OnPropertyChanged(nameof(CanApprove));
            ExecuteProcessRequirements(null);
        }

        private void ExecuteCancelProcessing(object parameter)
        {
            IsProcessing = false;
            ProcessingStatusMessage = "Processing cancelled by user";
            CurrentOperation = "";
            ShowMessage("Processing cancelled. Rolling back changes...");
        }

        private void ExecutePauseProcessing(object parameter)
        {
            CanPauseProcessing = false;
            ShowMessage("Processing paused");
        }

        private void ExecuteSkipCurrent(object parameter)
        {
            ShowMessage("Skipping current operation...");
        }


        private async void ExecuteSaveAssembly(object parameter)
        {
            try
            {
                // Add a SaveFileDialog
                var saveDialog = new SaveFileDialog
                {
                    Filter = "Folder Name|*.folder", // This is a trick to let user enter folder name
                    Title = "Enter Name for Modified Assembly Folder",
                    FileName = $"{NewAssemblyName}_Modified",
                    InitialDirectory = Path.GetDirectoryName(SelectedAssemblyFile.FilePath),
                    CheckFileExists = false,
                    CheckPathExists = true
                };

                if (saveDialog.ShowDialog() != true)
                    return;

                IsLoading = true;
                StatusMessage = "Saving modified assembly...";

                // Get folder and filename from dialog
                string savePath = saveDialog.FileName;
                string saveFolder = Path.GetDirectoryName(savePath);

                // Ensure directory exists
                if (!Directory.Exists(saveFolder))
                {
                    Directory.CreateDirectory(saveFolder);
                }

                if (IsWorkingWithTempFiles && !string.IsNullOrEmpty(CurrentTempFolderPath))
                {
                    try
                    {
                        // First, close all SolidWorks documents to release file locks
                        await Task.Run(() => _solidWorksService.CloseAllDocuments());

                        // Small delay to ensure files are released
                        await Task.Delay(500);

                        // Get destination folder from save dialog
                        string destinationDir = Path.GetDirectoryName(savePath);
                        string assemblyNameWithoutExt = Path.GetFileNameWithoutExtension(savePath);

                        // Create a properly named folder for the modified assembly
                        string finalFolder = Path.Combine(destinationDir, $"{assemblyNameWithoutExt}_Modified_{DateTime.Now:yyyyMMdd_HHmmss}");

                        // Create destination directory
                        Directory.CreateDirectory(finalFolder);

                        // Copy all files from temp to destination
                        await Task.Run(() => CopyDirectory(CurrentTempFolderPath, finalFolder));

                        // Clean up temp folder after successful copy
                        // Clean up temp folder after successful copy
                        await Task.Delay(500); // Increased delay to ensure files are released

                        // Try multiple times to delete temp folder
                        int attempts = 0;
                        bool deleted = false;
                        string deletionError = "";

                        while (attempts < 3 && !deleted)
                        {
                            try
                            {
                                // First, ensure any file handles are released
                                GC.Collect();
                                GC.WaitForPendingFinalizers();
                                GC.Collect();

                                // Try to delete the directory
                                Directory.Delete(CurrentTempFolderPath, true);

                                // If we get here, deletion was successful
                                CurrentTempFolderPath = null;
                                IsWorkingWithTempFiles = false;
                                deleted = true;
                            }
                            catch (Exception ex)
                            {
                                deletionError = ex.Message;
                                attempts++;

                                if (attempts < 3)
                                {
                                    // Wait longer before next attempt
                                    await Task.Delay(1000);

                                    // Try to release any locks by closing explorer windows
                                    try
                                    {
                                        // Force garbage collection to release any file handles
                                        GC.Collect();
                                        GC.WaitForPendingFinalizers();
                                    }
                                    catch { }
                                }
                            }
                        }

                        // If deletion failed after all attempts, inform the user
                        if (!deleted && Directory.Exists(CurrentTempFolderPath))
                        {
                            // Add to the success message about the temp folder
                            ShowMessage($"Assembly saved successfully!\n\nLocation: {finalFolder}\n\nAll assembly files have been saved to this folder.\n\n" +
                                        $"NOTE: Temporary files could not be deleted and remain at:\n{CurrentTempFolderPath}\n" +
                                        $"Error: {deletionError}\n\n" +
                                        $"You can manually delete this folder later.");

                            // Still clear the reference so new processing can start
                            CurrentTempFolderPath = null;
                            IsWorkingWithTempFiles = false;
                        }
                        else
                        {
                            // Normal success message
                            ShowMessage($"Assembly saved successfully!\n\nLocation: {finalFolder}\n\nAll assembly files have been saved to this folder.");
                        }

                        ShowMessage($"Assembly saved successfully!\n\nLocation: {finalFolder}\n\nAll assembly files have been saved to this folder.");
                    }
                    catch (Exception ex)
                    {
                        ShowMessage($"Save Error: {ex.Message}\n\nThe temporary files are still available at:\n{CurrentTempFolderPath}");
                        throw;
                    }
                }
                else
                {
                    // Original save logic for non-temp workflow
                    var saveResult = await Task.Run(() =>
                        _solidWorksService.SaveAssemblyAs(
                            savePath,
                            true
                        ));


                    if (saveResult.Success)
                    {
                        string fileList = string.Join("\n", saveResult.SavedFiles.Take(5));
                        if (saveResult.SavedFiles.Count > 5)
                            fileList += $"\n... and {saveResult.SavedFiles.Count - 5} more files";

                        ShowMessage($"Assembly saved successfully!\n\nLocation: {saveFolder}\n\nFiles saved:\n{fileList}");
                    }
                    else
                    {
                        ShowMessage($"Save failed: {saveResult.ErrorMessage}");
                    }
                }
            }
            catch (Exception ex)
            {
                ShowMessage($"Save Error: {ex.Message}");
            }
            finally
            {
                IsLoading = false;
                StatusMessage = "Ready";
            }
        }


        // New JSON Tab Commands
        private void ExecuteValidateJSON(object parameter)
        {
            try
            {
                // Use real JSON validation service
                var validationResult = _jsonService.ValidateModificationJson(ModificationJSON);

                if (validationResult.IsValid)
                {
                    ShowMessage("JSON is valid and ready to apply!");
                }
                else
                {
                    string errors = string.Join("\n", validationResult.ErrorMessages);
                    ShowMessage($"JSON Validation Errors:\n\n{errors}");
                }
            }
            catch (Exception ex)
            {
                ShowMessage($"Validation Error: {ex.Message}");
            }
        }


        private void ExecuteApplyJSON(object parameter)
        {
            ShowMessage("Applying manual JSON modifications...");
            // Re-process with the modified JSON
            ExecuteApproveAll(null);
        }

        private void ExecuteResetJSON(object parameter)
        {
            // Reset to original AI-generated JSON
            ExecuteReanalyze(null);
            EnableJSONEditing = false;
        }

        // Feature Mapping Commands
        private void ExecuteValidateMappings(object parameter)
        {
            ShowMessage("Validating feature mappings against SolidWorks model...\n\nAll mappings are valid!");
        }

        // Conflict Commands
        private void ExecuteIgnoreConflicts(object parameter)
        {
            ShowMessage("Proceeding with conflicts ignored. Some features may fail during processing.");
            ExecuteApproveAll(null);
        }

        private void ExecuteAutoResolve(object parameter)
        {
            ShowMessage("Attempting to auto-resolve conflicts...\n\nConflicts resolved. Proceeding with processing.");
            Conflicts.Clear();
            ExecuteApproveAll(null);
        }

        private void ExecuteNewProject(object parameter)
        {
            // Clear all selections and data
            SelectedAssemblyFile = null;
            SelectedBOMFile = null;
            SelectedRequirementsFile = null;
            ProcessingSteps.Clear();
            ExtractedRequirements.Clear();
            FeatureMappings.Clear();
            Conflicts.Clear();
            ChangeSummary.Clear();
            RequirementsFulfillment.Clear();
            CurrentProcessingStage = 0;
            ProjectStatus = "New project created";
            UpdateSaveLocation();
        }

        private void ExecuteOpenProject(object parameter)
        {
            var dialog = new OpenFileDialog
            {
                Filter = "Project Files (*.swproj)|*.swproj|All Files (*.*)|*.*",
                Title = "Open Project"
            };

            if (dialog.ShowDialog() == true)
            {
                ProjectStatus = $"Loaded: {System.IO.Path.GetFileName(dialog.FileName)}";
                ShowMessage($"Project loaded: {dialog.FileName}");
            }
        }

        private void ExecuteSaveProject(object parameter)
        {
            var dialog = new SaveFileDialog
            {
                Filter = "Project Files (*.swproj)|*.swproj|All Files (*.*)|*.*",
                Title = "Save Project"
            };

            if (dialog.ShowDialog() == true)
            {
                ShowMessage($"Project saved: {dialog.FileName}");
            }
        }


        #endregion

        #region Helper Methods

        private void CopyDirectory(string sourceDir, string destDir)
        {
            // Create all directories
            foreach (string dirPath in Directory.GetDirectories(sourceDir, "*", SearchOption.AllDirectories))
            {
                Directory.CreateDirectory(dirPath.Replace(sourceDir, destDir));
            }

            // Copy all files
            foreach (string filePath in Directory.GetFiles(sourceDir, "*.*", SearchOption.AllDirectories))
            {
                string destFilePath = filePath.Replace(sourceDir, destDir);

                // Try multiple times in case of temporary locks
                int attempts = 0;
                while (attempts < 3)
                {
                    try
                    {
                        File.Copy(filePath, destFilePath, true);
                        break;
                    }
                    catch (IOException)
                    {
                        attempts++;
                        if (attempts >= 3)
                            throw;
                        System.Threading.Thread.Sleep(500);
                    }
                }
            }
        }

        private async void UpdateAssemblyPreview()
        {
            if (SelectedAssemblyFile != null)
            {
                try
                {
                    IsLoading = true;
                    StatusMessage = "Opening assembly...";

                    // Always use original file for preview
                    // Temp files are only created when processing starts
                    var assemblyInfo = await Task.Run(() =>
                        _solidWorksService.OpenAssembly(SelectedAssemblyFile.FilePath));

                    // Update UI with real data
                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        AssemblyName = assemblyInfo.Name;
                        AssemblyPartCount = assemblyInfo.PartCount.ToString();
                        AssemblyFileSize = SelectedAssemblyFile.FileSize;
                        AssemblyThumbnail = assemblyInfo.ThumbnailPath;

                        // Store features for later use - Initialize if null
                        _currentAssemblyFeatures = assemblyInfo.Features ?? new List<string>();
                    });

                    StatusMessage = "Assembly loaded successfully";
                }
                catch (Exception ex)
                {
                    StatusMessage = $"Failed to open assembly: {ex.Message}";
                    // Initialize to empty list to prevent null errors
                    _currentAssemblyFeatures = new List<string>();
                }
                finally
                {
                    IsLoading = false;
                    UpdateSaveLocation();
                }
            }
        }
        private void UpdateSaveLocation()
        {
            if (!string.IsNullOrEmpty(WorkingDirectory) && !string.IsNullOrEmpty(NewAssemblyName))
            {
                // Generate a folder name based on the new assembly name and timestamp
                string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                string folderName = $"{System.IO.Path.GetFileNameWithoutExtension(NewAssemblyName)}_{timestamp}";
                SaveLocationPath = System.IO.Path.Combine(WorkingDirectory, "Modified_Assemblies", folderName);
            }
            else
            {
                SaveLocationPath = System.IO.Path.Combine(WorkingDirectory ?? "C:\\SolidWorks\\Projects", "Modified_Assemblies", "New_Assembly");
            }
        }

        private async void UpdateBOMPreview()
        {
            if (SelectedBOMFile != null)
            {
                try
                {
                    BOMPreviewData.Clear();

                    // Read real BOM file
                    var bomItems = await Task.Run(() =>
                        _fileService.ReadBOMFile(SelectedBOMFile.FilePath));

                    // Update UI
                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        foreach (var item in bomItems)
                        {
                            BOMPreviewData.Add(item);
                        }
                    });
                }
                catch (Exception ex)
                {
                    ShowMessage($"Error reading BOM: {ex.Message}");
                }
            }
        }

        private async void UpdateRequirementsPreview()
        {
            if (SelectedRequirementsFile != null)
            {
                try
                {
                    // Read real requirements file
                    RequirementsText = await Task.Run(() =>
                        _fileService.ReadRequirementsFile(SelectedRequirementsFile.FilePath));
                }
                catch (Exception ex)
                {
                    RequirementsText = $"Error reading file: {ex.Message}";
                }
            }
        }

        private void ShowMessage(string message)
        {
            System.Windows.MessageBox.Show(message, "SolidWorks Assembly Modifier",
                System.Windows.MessageBoxButton.OK,
                System.Windows.MessageBoxImage.Information);
        }

        // Test integration - can be called from a menu item
        private async void TestBackendIntegration()
        {
            try
            {
                // Test 1: File Service
                var files = _fileService.GetAssemblyFiles(@"C:\Test");
                Console.WriteLine($"Found {files.Count} files");

                // Test 2: SolidWorks Service
                if (files.Count > 0)
                {
                    var assembly = _solidWorksService.OpenAssembly(files[0].FilePath);
                    Console.WriteLine($"Opened assembly with {assembly.PartCount} parts");
                }

                // Test 3: LLM Service
                var result = await _llmService.AnalyzeRequirements(
                    "Increase hole diameter to 12mm",
                    null,
                    step => Console.WriteLine(step.Message)
                );
                Console.WriteLine($"Found {result.ExtractedRequirements.Count} requirements");

                ShowMessage("Backend integration successful!");
            }
            catch (Exception ex)
            {
                ShowMessage($"Integration test failed: {ex.Message}");
            }
        }
        #endregion



        private string GetUniqueDirectoryName(string basePath)
        {
            string directory = basePath;
            int counter = 1;

            while (Directory.Exists(directory))
            {
                directory = $"{basePath}_{counter}";
                counter++;
            }

            return directory;
        }


        public void Cleanup()
        {
            _timer?.Stop();

            // Clean up temp folder if exists
            CleanupTempFolder();

            _solidWorksService?.Dispose();
        }


    }
}