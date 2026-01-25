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
    /// 改进点：确保 Name 属性支持属性变更通知，以便在 UI 上的 TextBox 编辑时能正确同步。
    /// </summary>
    public class AnimPathItem : INotifyPropertyChanged
    {
        public ObjectId Id { get; set; }

        private string _name;
        public string Name
        {
            get => _name;
            set { _name = value; OnPropertyChanged(nameof(Name)); }
        }

        private int _groupNumber = 1;
        public int GroupNumber
        {
            get => _groupNumber;
            set { _groupNumber = value; OnPropertyChanged(nameof(GroupNumber)); }
        }

        private AnimLineStyle _lineStyle = AnimLineStyle.Solid;
        public AnimLineStyle LineStyle
        {
            get => _lineStyle;
            set { _lineStyle = value; OnPropertyChanged(nameof(LineStyle)); }
        }

        public System.Windows.Media.Color PathColor { get; set; }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}