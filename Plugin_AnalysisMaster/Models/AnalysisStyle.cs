using System;


namespace Plugin_AnalysisMaster.Models
{
    public enum PathCategory { None, Solid, Dashed, CustomPattern }

    // 定义所有必要的枚举
    public enum ArrowHeadType { None, Basic, SwallowTail, Circle, Square }
    public enum ArrowTailType { None, Swallow, Circle, Bar }
    public enum LineStyleType { Solid, DoubleLine, Dashed }

    public class AnalysisStyle
    {
        // 请将此属性添加到 AnalysisStyle 类中
        public double LineWeight { get; set; } = 0.30;
        // 几何参数
        public double ArrowSize { get; set; } = 8.0;
        public double ArrowAngle { get; set; } = 30.0;
        public double SwallowDepth { get; set; } = 0.4;

        // ✨ 新增/更新：中间路径的宽度参数
        public double StartWidth { get; set; } = 1.0; // 路径起始宽度
        public double EndWidth { get; set; } = 0.5;   // 路径结束宽度（用于渐变/渐细效果）

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
        public PathCategory PathType { get; set; } = PathCategory.Solid; // 路径分类

        // ✨ 新增：宽度控制

        public double MidWidth { get; set; } = 0.8; // 腰部宽度，实现两头宽中间窄的关键

        // ✨ 新增：线型控制
        public double LinetypeScale { get; set; } = 1.0;
    }
}