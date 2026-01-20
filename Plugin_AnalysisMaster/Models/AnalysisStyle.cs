using System;

namespace Plugin_AnalysisMaster.Models
{
    // ✨ 路径模式：无、连续线、阵列样式
    public enum PathCategory { None, Solid, Dashed, Pattern }
    public enum ArrowHeadType { None, Basic, SwallowTail, Circle, Square }
    public enum ArrowTailType { None, Swallow, Circle, Bar }

    /// <summary>
    /// 存储动线所有样式参数的数据模型。
    /// 新增 CapIndent 属性用于控制端头缩进，并保留起终点图块名称存储。
    /// </summary>
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

        // 3. 端头设置
        public string StartArrowType { get; set; } = "None"; // 存储真实块名 (如 Cap_Arrow1)
        public string EndArrowType { get; set; } = "None";   // 存储真实块名
        public double ArrowSize { get; set; } = 8.0;         // 对应缩放比例
        public double CapIndent { get; set; } = 0;          // ✨ 新增：起终点共用的缩进距离

        // 4. 路径模式配置
        public PathCategory PathType { get; set; } = PathCategory.Solid;
        public string SelectedBlockName { get; set; } = "";  // 存储真实块名 (如 Pat_Tree)
        public double PatternSpacing { get; set; } = 10.0;
        public double PatternScale { get; set; } = 1.0;
    }
}