using Autodesk.AutoCAD.Colors;
using System;
using System.Windows.Media;
using CadColor = Autodesk.AutoCAD.Colors.Color; // 为 CAD 颜色设置别名
using WpfColor = System.Windows.Media.Color;    // 为 WPF 颜色设置别名

namespace Plugin_AnalysisMaster.Models
{
    /// <summary>
    /// 头部/尾端样式枚举
    /// </summary>
    public enum ArrowHeadType
    {
        None,           // 无
        Basic,          // 普通三角形箭头
        SwallowTail,    // 燕尾箭头
        Circle,         // 实心圆点
        HollowCircle,   // 空心环
        Square          // 方块
    }

    /// <summary>
    /// 线条物理类型
    /// </summary>
    public enum LineStyleType
    {
        Solid,          // 实线
        Dashed,         // 短虚线
        Dotted,         // 点划线
        DoubleLine      // 双线 (外粗内细)
    }

    /// <summary>
    /// 分析线样式配置模型（参数化核心）
    /// </summary>
    public class AnalysisStyle
    {
        // --- 基础属性 ---
        public string Name { get; set; }                // 样式名称（如：深蓝动线）
        public System.Windows.Media.Color MainColor { get; set; }           // 主色调 (RGB)
        public double LineWeight { get; set; }          // 线宽

        // --- 几何特征 ---
        public LineStyleType LineType { get; set; }     // 线型
        public ArrowHeadType HeadType { get; set; }     // 箭头/起始端类型
        public ArrowHeadType TailType { get; set; }     // 末端类型

        // --- 详细参数 ---
        public double ArrowSize { get; set; } = 1.0;    // 箭头缩放比例
        public double SwallowDepth { get; set; } = 0.5; // 燕尾凹陷深度 (0.0 - 1.0)
        public bool IsCurved { get; set; } = false;     // 是否拟合为曲线
        public double Transparency { get; set; } = 0;   // 透明度 (0-255)

        // --- 图层控制 ---
        public string TargetLayer { get; set; } = "ANALYSIS_LINES"; // 目标图层
    }
}