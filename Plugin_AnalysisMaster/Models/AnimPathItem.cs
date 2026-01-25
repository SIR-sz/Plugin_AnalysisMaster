// 文件：AnimPathItem.cs

using Autodesk.AutoCAD.DatabaseServices;
using System.ComponentModel;

namespace Plugin_AnalysisMaster.Models
{
    /// <summary>
    /// 动画线型枚举：实线或虚线
    /// </summary>
    public enum AnimLineStyle
    {
        Solid = 0,
        Dash = 1
    }

    /// <summary>
    /// 动画路径条目模型
    /// </summary>
    public class AnimPathItem : INotifyPropertyChanged
    {
        public ObjectId Id { get; set; }
        public string Name { get; set; }

        private int _groupNumber = 1;
        public int GroupNumber
        {
            get => _groupNumber;
            set { _groupNumber = value; OnPropertyChanged(nameof(GroupNumber)); }
        }

        private AnimLineStyle _lineStyle = AnimLineStyle.Solid;
        /// <summary>
        /// 控制动画生长时的线型（实线/虚线）
        /// </summary>
        public AnimLineStyle LineStyle
        {
            get => _lineStyle;
            set { _lineStyle = value; OnPropertyChanged(nameof(LineStyle)); }
        }

        public System.Windows.Media.Color PathColor { get; set; }
        public double SamplingInterval { get; set; }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}