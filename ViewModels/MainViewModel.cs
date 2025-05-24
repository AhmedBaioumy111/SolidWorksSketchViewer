using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Input;
using System.Windows.Threading;
using Microsoft.Win32;
using SolidWorksSketchViewer.Helpers;
using SolidWorksSketchViewer.Models;
using SolidWorksSketchViewer.Services;

namespace SolidWorksSketchViewer.ViewModels
{
    public class MainViewModel : ViewModelBase
    {
        private DispatcherTimer _timer;
        private DateTime _startTime;
        private Random _random = new Random();

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
            set => SetProperty(ref _isLLMProcessing, value);
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
        public ICommand ValidateFilesCommand { get; private set; }
        public ICommand TestLLMCommand { get; private set; }
        public ICommand CheckSolidWorksCommand { get; private set; }
        public ICommand ShowDocumentationCommand { get; private set; }
        public ICommand ShowAPIStatusCommand { get; private set; }
        public ICommand ShowAboutCommand { get; private set; }
        public ICommand SettingsCommand { get; private set; }
        public ICommand HelpCommand { get; private set; }

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

        public MainViewModel()
        {
            // Initialize services
            _solidWorksService = new SolidWorksService();
            _fileService = new FileService();
            _llmService = new LLMService();
            _jsonService = new JsonProcessingService();

            InitializeCollections();
            InitializeCommands();
            InitializeMockData();
            StartTimer();
            UpdateSaveLocation();
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
            ValidateFilesCommand = new RelayCommand(ExecuteValidateFiles);
            TestLLMCommand = new RelayCommand(ExecuteTestLLM);
            CheckSolidWorksCommand = new RelayCommand(ExecuteCheckSolidWorks);
            ShowDocumentationCommand = new RelayCommand(_ => ShowMessage("Documentation would open here"));
            ShowAPIStatusCommand = new RelayCommand(_ => ShowMessage("API Status: All systems operational"));
            ShowAboutCommand = new RelayCommand(_ => ShowMessage("SolidWorks Assembly Modifier v1.0"));
            SettingsCommand = new RelayCommand(_ => ShowMessage("Settings dialog would open here"));
            HelpCommand = new RelayCommand(_ => ShowMessage("Help system would open here"));

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
            ExportPDFCommand = new RelayCommand(_ => ShowMessage("Exporting to PDF..."));
            ExportExcelCommand = new RelayCommand(_ => ShowMessage("Exporting to Excel..."));
            ExportLogCommand = new RelayCommand(_ => ShowMessage("Exporting processing log..."));
        }

        private void InitializeMockData()
        {
            // Add mock assembly files
            AssemblyFiles.Add(new FileItemModel
            {
                FileName = "Main_Assembly_v3.sldasm",
                FileSize = "25.4 MB",
                ValidationStatusColor = "Green"
            });
            AssemblyFiles.Add(new FileItemModel
            {
                FileName = "Sub_Assembly_Motor.sldasm",
                FileSize = "12.1 MB",
                ValidationStatusColor = "Green"
            });
            AssemblyFiles.Add(new FileItemModel
            {
                FileName = "Frame_Assembly.sldasm",
                FileSize = "8.7 MB",
                ValidationStatusColor = "Orange"
            });

            // Add mock BOM files
            BOMFiles.Add(new FileItemModel { FileName = "BOM_Main_Assembly.xlsx", FileSize = "124 KB" });
            BOMFiles.Add(new FileItemModel { FileName = "Parts_List_Rev2.csv", FileSize = "45 KB" });

            // Add mock requirements files
            RequirementsFiles.Add(new FileItemModel { FileName = "Customer_Requirements_v2.txt", FileSize = "12 KB" });
            RequirementsFiles.Add(new FileItemModel { FileName = "Design_Specs_2025.docx", FileSize = "256 KB" });
            RequirementsFiles.Add(new FileItemModel { FileName = "Dimension_Changes.json", FileSize = "4 KB" });
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

        private void ExecuteRefreshAssemblies(object parameter)
        {
            // Mock refresh - in real implementation, scan directory for files
            StatusMessage = "Refreshing file list...";
            // Add animation or actual file scanning here
            StatusMessage = "File list refreshed";
        }

        private async void ExecuteProcessRequirements(object parameter)
        {
            CurrentProcessingStage = 1; // Switch to AI Analysis tab
            IsLLMProcessing = true;
            LLMStatusMessage = "Analyzing Requirements with AI...";
            ProcessingSteps.Clear();

            // Simulate LLM processing steps
            await SimulateLLMProcessing();
        }

        private async Task SimulateLLMProcessing()
        {
            var steps = new[]
            {
                "Parsing requirements text...",
                "Extracting dimensions and constraints...",
                "Mapping to SolidWorks features...",
                "Validating modifications...",
                "Generating modification JSON..."
            };

            for (int i = 0; i < steps.Length; i++)
            {
                ProcessingSteps.Add(new ProcessingStep
                {
                    Status = "⏳",
                    Message = steps[i]
                });

                LLMProgress = (i + 1) * 20;
                await Task.Delay(1000);

                ProcessingSteps[i].Status = "✓";
            }

            // Generate mock extracted requirements
            ExtractedRequirements.Clear();
            ExtractedRequirements.Add(new ExtractedRequirement
            {
                Text = "Increase hole diameter from 10mm to 12mm",
                Type = "Dimension",
                Confidence = 95
            });
            ExtractedRequirements.Add(new ExtractedRequirement
            {
                Text = "Change material to Aluminum 6061",
                Type = "Material",
                Confidence = 88
            });
            ExtractedRequirements.Add(new ExtractedRequirement
            {
                Text = "Add 5mm chamfer to all edges",
                Type = "Feature",
                Confidence = 72
            });
            ExtractedRequirements.Add(new ExtractedRequirement
            {
                Text = "Ensure minimum wall thickness of 3mm",
                Type = "Constraint",
                Confidence = 45
            });

            // Generate mock feature mappings
            FeatureMappings.Clear();
            FeatureMappings.Add(new FeatureMapping
            {
                Requirement = "Increase hole diameter",
                TargetFeature = "Sketch3 - D1@Sketch3",
                CurrentValue = "10mm",
                NewValue = "12mm",
                Status = "Valid"
            });
            FeatureMappings.Add(new FeatureMapping
            {
                Requirement = "Change material",
                TargetFeature = "Part Properties",
                CurrentValue = "Steel",
                NewValue = "Aluminum 6061",
                Status = "Valid"
            });
            FeatureMappings.Add(new FeatureMapping
            {
                Requirement = "Add chamfer",
                TargetFeature = "Edge<1>, Edge<2>, Edge<3>",
                CurrentValue = "None",
                NewValue = "5mm Chamfer",
                Status = "Warning"
            });

            // Add a mock conflict
            Conflicts.Clear();
            Conflicts.Add(new ConflictItem
            {
                Title = "Potential Over-constraint",
                Description = "Adding 5mm chamfer to all edges may conflict with existing fillets on Edge<4> and Edge<5>",
                Resolution = "Consider excluding edges with existing features or reducing chamfer size"
            });

            // Generate mock JSON
            ModificationJSON = @"{
  ""modifications"": [
    {
      ""type"": ""dimension"",
      ""feature"": ""D1@Sketch3"",
      ""currentValue"": 10.0,
      ""newValue"": 12.0,
      ""units"": ""mm""
    },
    {
      ""type"": ""material"",
      ""property"": ""Material"",
      ""currentValue"": ""Steel"",
      ""newValue"": ""Aluminum 6061""
    },
    {
      ""type"": ""feature"",
      ""operation"": ""add_chamfer"",
      ""edges"": [""Edge<1>"", ""Edge<2>"", ""Edge<3>""],
      ""value"": 5.0,
      ""units"": ""mm""
    }
  ]
}";

            IsLLMProcessing = false;
            LLMStatusMessage = "Analysis complete. Review modifications below.";
        }

        private async void ExecuteApproveAll(object parameter)
        {
            CurrentProcessingStage = 2; // Switch to SolidWorks Processing tab
            await SimulateSolidWorksProcessing();
        }

        private async Task SimulateSolidWorksProcessing()
        {
            IsProcessing = true;
            CanPauseProcessing = true;
            ProcessingStatusMessage = "Modifying Assembly...";
            FeatureProcessingStatus.Clear();

            var features = new[]
            {
                ("Sketch3 - Dimension D1", true, ""),
                ("Material Properties", true, ""),
                ("Edge Chamfer Feature", false, "Overconstrained assembly"),
                ("Sketch5 - Dimension D2", true, ""),
                ("Mate3 - Distance", false, "Geometric conflict")
            };

            for (int i = 0; i < features.Length; i++)
            {
                var (name, success, error) = features[i];
                CurrentOperation = $"Processing: {name}";
                ProcessingProgress = (i + 1) * 20;

                var status = new FeatureProcessingStatus
                {
                    FeatureName = name,
                    StatusIcon = "⏳",
                    Message = "Processing...",
                    ProcessingTime = "0.0s",
                    BackgroundColor = "#FFF3E0"
                };

                FeatureProcessingStatus.Add(status);
                await Task.Delay(1500);

                // Update status after processing
                if (success)
                {
                    status.StatusIcon = "✓";
                    status.Message = "Successfully modified";
                    status.BackgroundColor = "#E8F5E9";
                }
                else
                {
                    status.StatusIcon = "❌";
                    status.Message = error;
                    status.BackgroundColor = "#FFEBEE";
                }
                status.ProcessingTime = $"{_random.Next(100, 2000) / 1000.0:F1}s";
            }

            IsProcessing = false;
            CanPauseProcessing = false;
            ProcessingStatusMessage = "Processing complete";
            CurrentOperation = "";

            // Update results
            UpdateResults();
        }

        private void UpdateResults()
        {
            SuccessSummary = "12 of 15 requirements successfully applied\nAssembly modified and saved";
            ProcessingTime = "Total processing time: 2 minutes 34 seconds";

            // Update change summary
            ChangeSummary.Clear();
            ChangeSummary.Add(new ChangeSummaryItem
            {
                Feature = "D1@Sketch3",
                OriginalValue = "10mm",
                NewValue = "12mm",
                Status = "✓"
            });
            ChangeSummary.Add(new ChangeSummaryItem
            {
                Feature = "Material",
                OriginalValue = "Steel",
                NewValue = "Al 6061",
                Status = "✓"
            });
            ChangeSummary.Add(new ChangeSummaryItem
            {
                Feature = "Chamfer1",
                OriginalValue = "None",
                NewValue = "Failed",
                Status = "❌"
            });

            // Update requirements fulfillment
            RequirementsFulfillment.Clear();
            RequirementsFulfillment.Add(new RequirementsFulfillmentItem
            {
                Requirement = "Increase hole diameter from 10mm to 12mm",
                StatusIcon = "✓",
                BackgroundColor = "#E8F5E9"
            });
            RequirementsFulfillment.Add(new RequirementsFulfillmentItem
            {
                Requirement = "Change material to Aluminum 6061",
                StatusIcon = "✓",
                BackgroundColor = "#E8F5E9"
            });
            RequirementsFulfillment.Add(new RequirementsFulfillmentItem
            {
                Requirement = "Add 5mm chamfer to all edges",
                StatusIcon = "❌",
                BackgroundColor = "#FFEBEE"
            });
            RequirementsFulfillment.Add(new RequirementsFulfillmentItem
            {
                Requirement = "Ensure minimum wall thickness of 3mm",
                StatusIcon = "⚠",
                BackgroundColor = "#FFF3E0"
            });
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

        private void ExecuteSaveAssembly(object parameter)
        {
            string message = $"Creating new folder and saving assembly as:\n{SaveLocationPath}\\{NewAssemblyName}.sldasm\n\nAll related files will be copied to the new folder.";
            ShowMessage(message);
        }

        // New JSON Tab Commands
        private void ExecuteValidateJSON(object parameter)
        {
            try
            {
                // In real implementation, parse and validate JSON
                // For now, just check if it's valid JSON syntax
                using (var doc = System.Text.Json.JsonDocument.Parse(ModificationJSON))
                {
                    ShowMessage("JSON is valid and ready to apply!");
                }
            }
            catch (Exception ex)
            {
                ShowMessage($"JSON Validation Error:\n{ex.Message}");
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

        private void ExecuteValidateFiles(object parameter)
        {
            ShowMessage("Validating all selected files...");
            // Update validation status colors
            foreach (var file in AssemblyFiles)
            {
                file.ValidationStatusColor = _random.Next(10) > 2 ? "Green" : "Orange";
            }
        }

        private async void ExecuteTestLLM(object parameter)
        {
            LLMStatus = "Testing...";
            LLMStatusColor = "Orange";
            await Task.Delay(2000);
            LLMStatus = "Connected";
            LLMStatusColor = "Green";
            ShowMessage("LLM API connection test successful!");
        }

        private async void ExecuteCheckSolidWorks(object parameter)
        {
            SolidWorksStatus = "Checking...";
            SolidWorksStatusColor = "Orange";
            await Task.Delay(1500);
            SolidWorksStatus = "Connected";
            SolidWorksStatusColor = "Green";
            ShowMessage("SolidWorks connection verified!");
        }
        #endregion

        #region Helper Methods
        private void UpdateAssemblyPreview()
        {
            if (SelectedAssemblyFile != null)
            {
                AssemblyName = SelectedAssemblyFile.FileName;
                AssemblyPartCount = _random.Next(10, 100).ToString();
                AssemblyFileSize = SelectedAssemblyFile.FileSize;
                // In real implementation, load actual thumbnail
                AssemblyThumbnail = null;
                UpdateSaveLocation();
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

        private void UpdateBOMPreview()
        {
            if (SelectedBOMFile != null)
            {
                BOMPreviewData.Clear();
                // Generate mock BOM data
                var parts = new[] { "Motor Mount", "Bearing Housing", "Shaft", "Cover Plate", "Fastener Set" };
                var materials = new[] { "Aluminum 6061", "Steel 1045", "Stainless 316", "ABS Plastic" };

                foreach (var part in parts)
                {
                    BOMPreviewData.Add(new BOMPreviewItem
                    {
                        PartName = part,
                        Quantity = _random.Next(1, 10),
                        Material = materials[_random.Next(materials.Length)],
                        Description = $"Component for main assembly"
                    });
                }
            }
        }

        private void UpdateRequirementsPreview()
        {
            if (SelectedRequirementsFile != null)
            {
                // Generate mock requirements text
                RequirementsText = @"DESIGN REQUIREMENTS - Rev 2.0
===============================

1. DIMENSIONAL CHANGES:
   - Increase main bore diameter from 10mm to 12mm
   - Adjust mounting hole spacing to 85mm centers
   - Reduce overall height by 15mm

2. MATERIAL SPECIFICATIONS:
   - Change base material from Steel to Aluminum 6061-T6
   - Apply anodized finish (black)
   
3. FEATURE MODIFICATIONS:
   - Add 5mm chamfer to all exposed edges
   - Include drainage groove (3mm wide, 2mm deep)
   - Add M6 threaded inserts at mounting points

4. CONSTRAINTS:
   - Maintain minimum wall thickness of 3mm
   - Ensure clearance of 2mm between moving parts
   - Weight must not exceed 1.2kg

5. ASSEMBLY REQUIREMENTS:
   - Parts must be assembled without special tools
   - Include alignment features for assembly
   - Maintain backward compatibility with Rev 1.0";
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

        #region Example: How to Replace Mock Methods with Real Implementation

        /* Example 1: Replace mock file refresh with real implementation
        private async void ExecuteRefreshAssemblies(object parameter)
        {
            try
            {
                StatusMessage = "Refreshing file list...";
                IsLoading = true;
                
                // Call real service instead of mock
                var files = await Task.Run(() => 
                    _fileService.GetAssemblyFiles(WorkingDirectory));
                
                AssemblyFiles.Clear();
                foreach (var file in files)
                {
                    AssemblyFiles.Add(file);
                }
                
                StatusMessage = $"Found {files.Count} assembly files";
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
        */

        /* Example 2: Replace mock LLM processing with real implementation
        private async void ExecuteProcessRequirements(object parameter)
        {
            CurrentProcessingStage = 1;
            IsLLMProcessing = true;
            LLMStatusMessage = "Analyzing Requirements with AI...";
            ProcessingSteps.Clear();

            try
            {
                // Get assembly context
                var assemblyContext = new AssemblyContext
                {
                    AssemblyName = SelectedAssemblyFile.FileName,
                    AvailableFeatures = _solidWorksService.GetFeatureList()
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

                // Update UI with results
                ExtractedRequirements = new ObservableCollection<ExtractedRequirement>(
                    result.ExtractedRequirements);
                FeatureMappings = new ObservableCollection<FeatureMapping>(
                    result.FeatureMappings);
                Conflicts = new ObservableCollection<ConflictItem>(
                    result.Conflicts);
                ModificationJSON = result.ModificationJSON;

                IsLLMProcessing = false;
                LLMStatusMessage = "Analysis complete. Review modifications below.";
            }
            catch (Exception ex)
            {
                IsLLMProcessing = false;
                LLMStatusMessage = $"Error: {ex.Message}";
                ShowMessage($"LLM Analysis failed: {ex.Message}");
            }
        }
        */

        #endregion

        public void Cleanup()
        {
            _timer?.Stop();
            _solidWorksService?.Dispose();
        }
    }
}