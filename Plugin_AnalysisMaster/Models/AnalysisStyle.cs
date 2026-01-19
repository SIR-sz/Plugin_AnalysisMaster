using System;

namespace Plugin_AnalysisMaster.Models
{
    // ✨ 统一枚举：确保包含 Pattern (自定义类)

    public enum PathCategory { None, Solid, Dashed, CustomPattern }
    public enum BuiltInPatternType { Dot, Square, ShortDash }
    public enum ArrowHeadType { None, Basic, SwallowTail, Circle, Square }
    public enum ArrowTailType { None, Swallow, Circle, Bar }
    public enum LineStyleType { Solid, DoubleLine, Dashed }

    public class AnalysisStyle
    {
        public double LineWeight { get; set; } = 0.30;
        public double ArrowSize { get; set; } = 8.0;
        public double StartWidth { get; set; } = 1.0;
        public double EndWidth { get; set; } = 0.5;
        public double MidWidth { get; set; } = 0.8;

        // ✨ 补全：AnalysisLineJig 依赖的头尾样式属性
        public ArrowHeadType StartCapStyle { get; set; } = ArrowHeadType.None;
        public ArrowHeadType EndCapStyle { get; set; } = ArrowHeadType.Basic;
        public ArrowHeadType HeadType { get; set; } = ArrowHeadType.Basic; // 对应旧代码
        public ArrowTailType TailType { get; set; } = ArrowTailType.None;

        public System.Windows.Media.Color MainColor { get; set; } = System.Windows.Media.Colors.SteelBlue;
        public double Transparency { get; set; } = 0;
        public bool IsCurved { get; set; } = true;

        public string TargetLayer { get; set; } = "ANALYSIS_LINES";
        public PathCategory PathType { get; set; } = PathCategory.Solid;

        // ✨ 补全：线型与阵列参数
        public double LinetypeScale { get; set; } = 1.0;
        public string SelectedLinetype { get; set; } = "DASHED";
        public BuiltInPatternType SelectedBuiltInPattern { get; set; } = BuiltInPatternType.ShortDash;
        public string CustomBlockName { get; set; } = "";
        public double SwallowDepth { get; set; } = 0.4;
    }
}