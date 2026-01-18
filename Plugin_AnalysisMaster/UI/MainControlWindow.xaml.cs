using System;
using System.Windows;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Plugin_AnalysisMaster.Models;
using Plugin_AnalysisMaster.Services;

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

            // 1. 根据点击的按钮生成样式参数 (这里可以根据 Tag 区分)
            AnalysisStyle selectedStyle = new AnalysisStyle
            {
                ArrowSize = SizeSlider.Value,
                MainColor = ((System.Windows.Media.SolidColorBrush)btn.Background).Color
            };

            if (btn.Tag.ToString().Contains("Swallow"))
                selectedStyle.HeadType = ArrowHeadType.SwallowTail;
            else
                selectedStyle.HeadType = ArrowHeadType.Basic;

            // 2. 将焦点转回 CAD 并开始取点
            DrawLineInCad(selectedStyle);
        }

        private async void DrawLineInCad(AnalysisStyle style)
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            // 暂时隐藏窗口，避免遮挡
            this.Hide();

            try
            {
                PromptPointOptions ppo1 = new PromptPointOptions("\n请选择起点: ");
                PromptPointResult ppr1 = ed.GetPoint(ppo1);
                if (ppr1.Status != PromptStatus.OK) return;

                PromptPointOptions ppo2 = new PromptPointOptions("\n请选择终点: ");
                ppo2.UseBasePoint = true;
                ppo2.BasePoint = ppr1.Value;
                PromptPointResult ppr2 = ed.GetPoint(ppo2);
                if (ppr2.Status != PromptStatus.OK) return;

                // 调用之前定义的 Service 层绘图引擎
                GeometryEngine.DrawAnalysisLine(ppr1.Value, ppr2.Value, style);
            }
            finally
            {
                // 绘图结束，重新显示窗口
                this.Show();
            }
        }
    }
}