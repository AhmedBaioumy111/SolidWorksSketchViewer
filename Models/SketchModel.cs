using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace SolidWorksSketchViewer.Models
{
    public class SketchModel
    {
        public string Name { get; set; }
        public string Id { get; set; }
        public bool IsActive { get; set; }
        public ObservableCollection<DimensionModel> Dimensions { get; set; }
        public string FeatureType { get; set; }
        public bool IsVisible { get; set; }
        public int EntityCount { get; set; }

        public SketchModel()
        {
            Dimensions = new ObservableCollection<DimensionModel>();
        }
    }

    public class DimensionModel
    {
        public string Name { get; set; }
        public string Value { get; set; }
        public string Type { get; set; }
        public bool IsReference { get; set; }
        public string Units { get; set; }
    }

    public class SolidWorksDocumentModel
    {
        public string FilePath { get; set; }
        public string FileName { get; set; }
        public string DocumentType { get; set; }
        public ObservableCollection<SketchModel> Sketches { get; set; }

        public SolidWorksDocumentModel()
        {
            Sketches = new ObservableCollection<SketchModel>();
        }
    }
}