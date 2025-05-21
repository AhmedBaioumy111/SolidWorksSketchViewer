using System;
using System.Windows;
using System.Windows.Input;
using System.Collections.ObjectModel;
using Microsoft.Win32;
using SolidWorksSketchViewer.Helpers;
using SolidWorksSketchViewer.Models;

namespace SolidWorksSketchViewer.ViewModels
{
    public class MainViewModel : ViewModelBase
    {
        private SolidWorksService _solidWorksService;
        private SolidWorksDocumentModel _currentDocument;
        private SketchModel _selectedSketch;
        private string _filePath;
        private bool _isLoading;
        private string _statusMessage;

        public MainViewModel()
        {
            // Initialize the SolidWorks service
            _solidWorksService = new SolidWorksService();

            // Initialize commands
            BrowseFileCommand = new RelayCommand(ExecuteBrowseFile);
            SelectSketchCommand = new RelayCommand(ExecuteSelectSketch, CanExecuteSelectSketch);

            // Set default values
            StatusMessage = "Ready. Please select a SOLIDWORKS file.";
            IsLoading = false;
        }

        public SolidWorksDocumentModel CurrentDocument
        {
            get => _currentDocument;
            set => SetProperty(ref _currentDocument, value);
        }

        public SketchModel SelectedSketch
        {
            get => _selectedSketch;
            set
            {
                if (SetProperty(ref _selectedSketch, value) && value != null)
                {
                    // Select the sketch in SOLIDWORKS when a sketch is selected in the UI
                    _solidWorksService.SelectSketch(value.Id);
                }
            }
        }

        public string FilePath
        {
            get => _filePath;
            set => SetProperty(ref _filePath, value);
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

        public ICommand BrowseFileCommand { get; }
        public ICommand SelectSketchCommand { get; }

        private void ExecuteBrowseFile(object parameter)
        {
            try
            {
                var openFileDialog = new OpenFileDialog
                {
                    Filter = "SOLIDWORKS Files (*.sldprt;*.sldasm)|*.sldprt;*.sldasm|All Files (*.*)|*.*",
                    Title = "Select a SOLIDWORKS Part or Assembly"
                };

                if (openFileDialog.ShowDialog() == true)
                {
                    // Update file path and notify UI
                    FilePath = openFileDialog.FileName;
                    StatusMessage = "Loading file...";
                    IsLoading = true;

                    // Load the document in SOLIDWORKS
                    Application.Current.Dispatcher.BeginInvoke(new Action(() =>
                    {
                        try
                        {
                            CurrentDocument = _solidWorksService.OpenDocument(FilePath);

                            if (CurrentDocument != null)
                            {
                                StatusMessage = $"File loaded: {CurrentDocument.FileName} - Found {CurrentDocument.Sketches.Count} sketches.";
                            }
                            else
                            {
                                StatusMessage = "Failed to load the document.";
                            }
                        }
                        catch (Exception ex)
                        {
                            StatusMessage = $"Error: {ex.Message}";
                        }
                        finally
                        {
                            IsLoading = false;
                        }
                    }));
                }
            }
            catch (Exception ex)
            {
                StatusMessage = $"Error: {ex.Message}";
                IsLoading = false;
            }
        }

        private void ExecuteSelectSketch(object parameter)
        {
            var sketch = parameter as SketchModel;
            if (sketch != null)
            {
                SelectedSketch = sketch;
                StatusMessage = $"Selected sketch: {sketch.Name}";
            }
        }

        private bool CanExecuteSelectSketch(object parameter)
        {
            return parameter is SketchModel && !IsLoading;
        }

        // Clean up resources
        public void Cleanup()
        {
            _solidWorksService?.Dispose();
        }
    }
}