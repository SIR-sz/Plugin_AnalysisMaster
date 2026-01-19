using System;

namespace Plugin_AnalysisMaster.Models
{
    // ✨ 路径模式：无、连续线、阵列样式
    public enum PathCategory { None, Solid, Dashed, Pattern }
    public enum ArrowHeadType { None, Basic, SwallowTail, Circle, Square }
    public enum ArrowTailType { None, Swallow, Circle, Bar }

    public class AnalysisStyle
    {
        // 1. 基础属性
        public string TargetLayer { get; set; } = "ANALYSIS_LINES";
        public System.Windows.Media.Color MainColor { get; set; } = System.Windows.Media.Colors.SteelBlue;
        public double Transparency { get; set; } = 0;
        public bool IsCurved { get; set; } = true;

        // 2. 几何宽度 (Solid 模式渐变)
        public double StartWidth { get; set; } = 1.0;
        public double MidWidth { get; set; } = 0.8;
        public double EndWidth { get; set; } = 0.5;

        // 3. 端头设置 (补全以修复 CS1061)
        public ArrowHeadType StartCapStyle { get; set; } = ArrowHeadType.None;
        public ArrowHeadType EndCapStyle { get; set; } = ArrowHeadType.Basic;
        public ArrowHeadType HeadType { get; set; } = ArrowHeadType.Basic; // 供 Jig 使用
        public ArrowTailType TailType { get; set; } = ArrowTailType.None;
        public double ArrowSize { get; set; } = 8.0;
        public double SwallowDepth { get; set; } = 0.4;

        // 4. 阵列样式配置
        public PathCategory PathType { get; set; } = PathCategory.Solid;
        public string SelectedBlockName { get; set; } = "";    // 选中的块名
        public string CustomBlockName { get; set; } = "";      // 兼容旧代码
        public double PatternSpacing { get; set; } = 10.0;     // 阵列间距
        public double PatternScale { get; set; } = 1.0;       // 整体缩放
        public double LinetypeScale { get; set; } = 1.0;      // 兼容旧代码
        public string SelectedLinetype { get; set; } = "CONTINUOUS";
    }
}