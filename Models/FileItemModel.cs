using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace SolidWorksSketchViewer.Models
{
    // Base class for all file items
    public class FileItemModel : INotifyPropertyChanged
    {
        private string _validationStatusColor = "Green";

        public string FileName { get; set; }
        public string FilePath { get; set; }
        public string FileSize { get; set; }
        public DateTime LastModified { get; set; }


        public string ValidationStatusColor
        {
            get => _validationStatusColor;
            set
            {
                _validationStatusColor = value;
                OnPropertyChanged();
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    // BOM Preview Data Model
    public class BOMPreviewItem
    {
        public string PartName { get; set; }
        public int Quantity { get; set; }
        public string Material { get; set; }
        public string Description { get; set; }
    }

    // Processing Step Model
    public class ProcessingStep
    {
        public string Status { get; set; } // ✓, ⚠, ❌, ⏳
        public string Message { get; set; }
    }

    // Extracted Requirement Model
    public class ExtractedRequirement
    {
        public string Text { get; set; }
        public string Type { get; set; } // Dimension, Constraint, Material, etc.
        public double Confidence { get; set; }
        public bool IsHighConfidence => Confidence >= 80;
        public bool IsLowConfidence => Confidence < 50;
    }

    // Feature Mapping Model
    public class FeatureMapping
    {
        public string Requirement { get; set; }
        public string TargetFeature { get; set; }
        public string CurrentValue { get; set; }
        public string NewValue { get; set; }
        public string Status { get; set; } // Valid, Warning, Error
    }

    // Conflict Model
    public class ConflictItem
    {
        public string Title { get; set; }
        public string Description { get; set; }
        public string Resolution { get; set; }
    }

    // Feature Processing Status Model
    public class FeatureProcessingStatus
    {
        public string FeatureName { get; set; }
        public string StatusIcon { get; set; } // ✓, ⚠, ❌, ⏳
        public string Message { get; set; }
        public string ProcessingTime { get; set; }
        public string BackgroundColor { get; set; }
    }

    // Change Summary Item
    public class ChangeSummaryItem
    {
        public string Feature { get; set; }
        public string OriginalValue { get; set; }
        public string NewValue { get; set; }
        public string Status { get; set; }
    }

    // Requirements Fulfillment Item
    public class RequirementsFulfillmentItem
    {
        public string Requirement { get; set; }
        public string StatusIcon { get; set; } // ✓, ⚠, ❌
        public string BackgroundColor { get; set; }
    }
}