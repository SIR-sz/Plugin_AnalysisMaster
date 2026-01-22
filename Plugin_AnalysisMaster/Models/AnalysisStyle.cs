using System;

namespace Plugin_AnalysisMaster.Models
{
    // ✨ 路径模式：无、连续线、阵列样式
    public enum PathCategory { None, Solid, Dashed, Pattern }
    public enum ArrowHeadType { None, Basic, SwallowTail, Circle, Square }
    public enum ArrowTailType { None, Swallow, Circle, Bar }

    /// <summary>
    /// 存储动线所有样式参数的数据模型。
    /// 修改内容：新增了 SamplingInterval 属性，用于控制动画路径点的记录密度。
    /// </summary>
    public class AnalysisStyle
    {
        public string TargetLayer { get; set; } = "ANALYSIS_LINES";
        public System.Windows.Media.Color MainColor { get; set; } = System.Windows.Media.Colors.SteelBlue;
        public bool IsCurved { get; set; } = true;

        // 几何宽度
        public double StartWidth { get; set; } = 1.0;
        public double MidWidth { get; set; } = 1.0;
        public double EndWidth { get; set; } = 1.0;

        // 端头设置
        public string StartArrowType { get; set; } = "None";
        public string EndArrowType { get; set; } = "None";
        public double ArrowSize { get; set; } = 1.0;
        public double CapIndent { get; set; } = 0;

        // 路径模式配置
        public PathCategory PathType { get; set; } = PathCategory.Solid;
        public string SelectedBlockName { get; set; } = "";

        // 组合模式支持
        public bool IsComposite { get; set; } = false;
        public string SelectedBlockName2 { get; set; } = "";

        public double PatternSpacing { get; set; } = 2.0;
        public double PatternScale { get; set; } = 1.0;

        // 动画控制
        public bool IsAnimated { get; set; } = false;
        public int AnimationDelay { get; set; } = 10;

        // ✨ 新增：采样间距（用于存储和回放）
        public double SamplingInterval { get; set; } = 5.0;
    }
}