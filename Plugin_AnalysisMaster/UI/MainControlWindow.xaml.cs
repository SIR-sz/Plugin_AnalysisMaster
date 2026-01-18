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

        public MainControlWindow()
        {
            InitializeComponent();
            // 初始设置
            StartCapCombo.SelectedIndex = 0;
            SegmentStyleCombo.SelectedIndex = 0;
            EndCapCombo.SelectedIndex = 0;
            UpdatePreview();
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
        private void UpdatePreview()
        {
            if (PreviewCanvas == null) return;
            PreviewCanvas.Children.Clear();

            var brush = new SolidColorBrush(_currentStyle.MainColor);
            double centerY = PreviewCanvas.Height / 2;
            double startX = 50;
            double endX = 250;
            double s = _currentStyle.ArrowSize;

            // 1. 绘制中间路径 (根据 StartWidth 和 EndWidth 模拟渐变效果)
            // 预览中使用多边形来模拟带有宽度的路径
            Polygon body = new Polygon { Fill = brush, Opacity = 0.8 };
            double w1 = _currentStyle.StartWidth * 5; // 放大预览比例
            double w2 = _currentStyle.EndWidth * 5;

            body.Points.Add(new Point(startX, centerY - w1));
            body.Points.Add(new Point(endX, centerY - w2));
            body.Points.Add(new Point(endX, centerY + w2));
            body.Points.Add(new Point(startX, centerY + w1));
            PreviewCanvas.Children.Add(body);

            // 2. 绘制起点端头 (根据 StartCapStyle)
            if (_currentStyle.StartCapStyle == ArrowHeadType.Circle)
            {
                Ellipse dot = new Ellipse { Fill = brush, Width = s, Height = s };
                Canvas.SetLeft(dot, startX - s / 2);
                Canvas.SetTop(dot, centerY - s / 2);
                PreviewCanvas.Children.Add(dot);
            }

            // 3. 绘制终点端头 (根据 EndCapStyle)
            if (_currentStyle.EndCapStyle != ArrowHeadType.None)
            {
                Polygon head = new Polygon { Fill = brush };
                if (_currentStyle.EndCapStyle == ArrowHeadType.SwallowTail)
                {
                    // 燕尾预览算法
                    head.Points.Add(new Point(endX + s, centerY));
                    head.Points.Add(new Point(endX, centerY - s / 2));
                    head.Points.Add(new Point(endX + s * 0.4, centerY));
                    head.Points.Add(new Point(endX, centerY + s / 2));
                }
                else // 标准三角形
                {
                    head.Points.Add(new Point(endX + s, centerY));
                    head.Points.Add(new Point(endX, centerY - s * 0.4));
                    head.Points.Add(new Point(endX, centerY + s * 0.4));
                }
                PreviewCanvas.Children.Add(head);
            }
        }

        // ✨ 修复 NullReferenceException 并实现三段式参数同步
        private void SyncStyleFromUI()
        {
            // 1. 确保样式对象已实例化
            if (_currentStyle == null) _currentStyle = new AnalysisStyle();

            // 2. 安全检查：如果控件尚未加载，直接返回，避免初始化时的空引用崩溃
            if (SizeSlider == null || WeightSlider == null || StartCapCombo == null || EndCapCombo == null)
                return;

            // 3. 同步几何参数
            _currentStyle.ArrowSize = SizeSlider.Value;
            _currentStyle.LineWeight = WeightSlider.Value;

            // 4. 实现解构逻辑：独立获取起点和终点的端头样式
            // 映射 ComboBox 的索引到枚举
            if (StartCapCombo.SelectedIndex != -1)
                _currentStyle.StartCapStyle = (ArrowHeadType)StartCapCombo.SelectedIndex;

            if (EndCapCombo.SelectedIndex != -1)
                _currentStyle.EndCapStyle = (ArrowHeadType)EndCapCombo.SelectedIndex;

            // 5. 路径样式处理
            if (SegmentStyleCombo != null && SegmentStyleCombo.SelectedIndex != -1)
            {
                _currentStyle.LineType = (LineStyleType)SegmentStyleCombo.SelectedIndex;
                // 如果是“渐细线”，设置不同的前后宽度
                if (SegmentStyleCombo.SelectedIndex == 1) // 假设索引1是渐细线
                {
                    _currentStyle.StartWidth = _currentStyle.LineWeight * 1.5;
                    _currentStyle.EndWidth = _currentStyle.LineWeight * 0.5;
                }
                else
                {
                    _currentStyle.StartWidth = _currentStyle.LineWeight;
                    _currentStyle.EndWidth = _currentStyle.LineWeight;
                }
            }
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