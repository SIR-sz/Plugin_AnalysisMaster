using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Plugin_AnalysisMaster.Models;
using Plugin_AnalysisMaster.Services;
using System;
using System.Windows;


namespace Plugin_AnalysisMaster.UI
{
    public partial class MainControlWindow : Window
    {
        // 单例模式，防止重复打开窗口
        private static MainControlWindow _instance;

        public MainControlWindow()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 供 MainTool 调用的静态启动方法
        /// </summary>
        public static void ShowTool()
        {
            if (_instance == null)
            {
                _instance = new MainControlWindow();
                _instance.Closed += (s, e) => _instance = null;
                Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessWindow(_instance);
            }
            else
            {
                _instance.Focus();
            }
        }

        private void StyleButton_Click(object sender, RoutedEventArgs e)
        {
            var btn = sender as System.Windows.Controls.Button;
            if (btn == null) return;

            AnalysisStyle selectedStyle = new AnalysisStyle
            {
                ArrowSize = SizeSlider.Value,
                MainColor = ((System.Windows.Media.SolidColorBrush)btn.Background).Color,
                // ✨ 读取 UI 上的曲线开关状态
                IsCurved = CurveCheckBox.IsChecked ?? false,
                LineWeight = 0.30 // 默认中等线宽
            };

            if (btn.Tag.ToString().Contains("Swallow"))
                selectedStyle.HeadType = ArrowHeadType.SwallowTail;
            else
                selectedStyle.HeadType = ArrowHeadType.Basic;

            DrawLineInCad(selectedStyle);
        }

        private void DrawLineInCad(AnalysisStyle style)
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            this.Hide();

            try
            {
                // 1. 获取起始点
                PromptPointResult ppr = ed.GetPoint("\n请选择动线起点: ");
                if (ppr.Status != PromptStatus.OK) return;

                // 2. 初始化 Jig
                AnalysisLineJig jig = new AnalysisLineJig(ppr.Value, style);

                // 3. 循环交互
                while (true)
                {
                    PromptResult res = ed.Drag(jig);

                    if (res.Status == PromptStatus.OK)
                    {
                        // 关键：将采样到的临时点正式加入点集
                        // 这里直接使用您在 Jig 中采样的结果
                        // 假设您在 Jig 类中暴露了最后采样点的属性：jig.LastPoint
                        jig.GetPoints().Add(jig.LastPoint);
                    }
                    else if (res.Status == PromptStatus.None) // 回车结束
                    {
                        if (jig.GetPoints().Count >= 2)
                            GeometryEngine.DrawAnalysisLine(jig.GetPoints(), style);
                        break;
                    }
                    else break;
                }
            }
            finally
            {
                this.Show();
            }
        }

    }
}