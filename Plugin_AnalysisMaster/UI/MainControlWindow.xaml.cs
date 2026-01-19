using Autodesk.AutoCAD.EditorInput;
using Plugin_AnalysisMaster.Models;
using Plugin_AnalysisMaster.Services;
using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Shapes;
using AcApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace Plugin_AnalysisMaster.UI
{
    public partial class MainControlWindow : Window
    {
        private AnalysisStyle _currentStyle = new AnalysisStyle();
        private static MainControlWindow _instance;

        public static void ShowTool()
        {
            if (_instance == null || !_instance.IsLoaded)
            {
                _instance = new MainControlWindow();
                AcApp.ShowModelessWindow(_instance);
            }
            else { _instance.Activate(); }
        }

        // ✨ 修改后的构造函数
        public MainControlWindow()
        {
            InitializeComponent();

            // 1. 初始化模型实例
            _currentStyle = new AnalysisStyle();

            // 2. 核心修复：注册 Loaded 事件，在窗口完全加载后执行第一次同步
            this.Loaded += (s, e) =>
            {
                SyncStyleFromUI(); // 确保 Slider 的初始值（如 1.0）被同步到模型
                UpdatePreview();   // 更新预览图
            };
        }

        // ✨ 核心方法：参数改变触发预览更新
        // ✨ 增加加载状态检查，彻底解决初始化崩溃问题
        private void OnParamChanged(object sender, EventArgs e)
        {
            // 关键：如果窗口还没加载完成，不要执行逻辑
            if (!this.IsLoaded) return;

            SyncStyleFromUI();
            UpdatePreview();
        }

        private void Color_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn)
            {
                _currentStyle.MainColor = ((SolidColorBrush)btn.Background).Color;
                UpdatePreview();
            }
        }

        // ✨ 商业化预览逻辑：在 Canvas 上模拟绘制动线
        // ✨ 商业级实时预览逻辑：完整模拟[起点+中间+终点]的组合效果
        // ✨ 完整的预览方法：在画布上绘制具有渐变宽度的“皮肤”线
        private void UpdatePreview()
        {
            if (PreviewCanvas == null || _currentStyle == null) return;
            PreviewCanvas.Children.Clear();

            var brush = new SolidColorBrush(_currentStyle.MainColor);
            double centerY = PreviewCanvas.Height / 2;
            double startX = 60;
            double endX = 240;

            // 1. 绘制主体路径预览 (模拟骨架与皮肤的关系)
            // 使用 Polygon 模拟从 StartWidth 到 EndWidth 的物理变化
            Polygon body = new Polygon { Fill = brush, Opacity = 0.8 };

            // 放大比例以在预览中可见 (4-5倍较合适)
            double w1 = _currentStyle.StartWidth * 4;
            double w2 = _currentStyle.EndWidth * 4;

            body.Points.Add(new Point(startX, centerY - w1));
            body.Points.Add(new Point(endX, centerY - w2));
            body.Points.Add(new Point(endX, centerY + w2));
            body.Points.Add(new Point(startX, centerY + w1));

            PreviewCanvas.Children.Add(body);

            // 2. 绘制终点端头 (End Cap)
            double s = _currentStyle.ArrowSize;
            if (_currentStyle.EndCapStyle != ArrowHeadType.None)
            {
                Polygon head = new Polygon { Fill = brush };
                head.Points.Add(new Point(endX + s, centerY));
                head.Points.Add(new Point(endX, centerY - s * 0.4));
                head.Points.Add(new Point(endX, centerY + s * 0.4));
                PreviewCanvas.Children.Add(head);
            }

            // 3. 绘制起点端头 (Start Cap)
            if (_currentStyle.StartCapStyle == ArrowHeadType.Circle)
            {
                Ellipse dot = new Ellipse { Fill = brush, Width = 8, Height = 8 };
                Canvas.SetLeft(dot, startX - 4);
                Canvas.SetTop(dot, centerY - 4);
                PreviewCanvas.Children.Add(dot);
            }
        }

        // ✨ 修复 NullReferenceException 并实现三段式参数同步
        // ✨ 核心逻辑：将 UI 上的起点、终点宽度同步到模型
        // ✨ 完整的同步方法：支持三组参数并修复透明度引用
        private void SyncStyleFromUI()
        {
            if (_currentStyle == null) _currentStyle = new AnalysisStyle();

            // 严谨的空检查：确保所有涉及到的 UI 控件都已实例化
            if (SizeSlider == null || StartWidthSlider == null || EndWidthSlider == null ||
                TransSlider == null || StartCapCombo == null || EndCapCombo == null)
                return;

            // 1. 同步端头尺寸
            _currentStyle.ArrowSize = SizeSlider.Value;

            // 2. 同步起点/终点物理宽度
            _currentStyle.StartWidth = StartWidthSlider.Value;
            _currentStyle.EndWidth = EndWidthSlider.Value;

            // 3. 同步组合样式 (起点端头、路径类型、终点端头)
            if (StartCapCombo.SelectedIndex != -1)
                _currentStyle.StartCapStyle = (ArrowHeadType)StartCapCombo.SelectedIndex;

            if (EndCapCombo.SelectedIndex != -1)
                _currentStyle.EndCapStyle = (ArrowHeadType)EndCapCombo.SelectedIndex;

            if (SegmentStyleCombo != null && SegmentStyleCombo.SelectedIndex != -1)
                _currentStyle.LineType = (LineStyleType)SegmentStyleCombo.SelectedIndex;

            // 4. 同步透明度 (修复 CS0103)
            _currentStyle.Transparency = TransSlider.Value;
        }

        private void StartDraw_Click(object sender, RoutedEventArgs e)
        {
            this.Hide();
            try
            {
                var doc = AcApp.DocumentManager.MdiActiveDocument;
                var ppr = doc.Editor.GetPoint("\n指定起点: ");
                if (ppr.Status == PromptStatus.OK)
                {
                    AnalysisLineJig jig = new AnalysisLineJig(ppr.Value, _currentStyle);
                    while (doc.Editor.Drag(jig).Status == PromptStatus.OK)
                    {
                        jig.GetPoints().Add(jig.LastPoint);
                    }
                    if (jig.GetPoints().Count >= 2)
                        GeometryEngine.DrawAnalysisLine(jig.GetPoints(), _currentStyle);
                }
            }
            finally { this.Show(); }
        }

        private void Close_Click(object sender, RoutedEventArgs e) => this.Close();
    }
}