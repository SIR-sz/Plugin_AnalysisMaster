using System;


namespace Plugin_AnalysisMaster.Models
{

    // 定义所有必要的枚举
    public enum ArrowHeadType { None, Basic, SwallowTail, Circle, Square }
    public enum ArrowTailType { None, Swallow, Circle, Bar }
    public enum LineStyleType { Solid, DoubleLine, Dashed }

    public class AnalysisStyle
    {
        // 请将此属性添加到 AnalysisStyle 类中
        public double LineWeight { get; set; } = 0.30;
        // 几何参数
        public double ArrowSize { get; set; } = 5.0;
        public double ArrowAngle { get; set; } = 30.0;
        public double SwallowDepth { get; set; } = 0.4;

        // ✨ 新增/更新：中间路径的宽度参数
        public double StartWidth { get; set; } = 0.3; // 路径起始宽度
        public double EndWidth { get; set; } = 0.3;   // 路径结束宽度（用于渐变/渐细效果）

        // ✨ 确保头尾枚举是独立的
        public ArrowHeadType StartCapStyle { get; set; } = ArrowHeadType.None;
        public ArrowHeadType EndCapStyle { get; set; } = ArrowHeadType.Basic;

        // 颜色与表现
        public System.Windows.Media.Color MainColor { get; set; } = System.Windows.Media.Colors.SteelBlue;
        public double Transparency { get; set; } = 0; // 0-255
        public bool IsCurved { get; set; } = true;

        // 类型控制
        public ArrowHeadType HeadType { get; set; } = ArrowHeadType.SwallowTail;
        public ArrowTailType TailType { get; set; } = ArrowTailType.None;
        public LineStyleType LineType { get; set; } = LineStyleType.Solid;

        // 图层
        public string TargetLayer { get; set; } = "ANALYSIS_LINES";
    }
}