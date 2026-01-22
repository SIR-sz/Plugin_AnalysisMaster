using Autodesk.AutoCAD.DatabaseServices;
using System;
using System.ComponentModel;

namespace Plugin_AnalysisMaster.Models
{
    public class AnimPathItem : INotifyPropertyChanged
    {
        public ObjectId Id { get; set; }
        public string Name { get; set; } // 比如 "图层: 道路 - 长度: 150"

        private int _groupNumber = 1;
        public int GroupNumber
        {
            get => _groupNumber;
            set { _groupNumber = value; OnPropertyChanged(nameof(GroupNumber)); }
        }

        public System.Windows.Media.Color PathColor { get; set; }
        public double SamplingInterval { get; set; }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}