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
        // 请替换完整的构造函数
        public MainControlWindow()
        {
            InitializeComponent();

            _currentStyle = new AnalysisStyle();

            // 注册加载事件进行初始同步
            this.Loaded += (s, e) =>
            {
                SyncStyleFromUI();
                UpdatePreview();
            };

            // ✨ 移除对 SegmentStyleCombo 的引用，保留现有的初始化
            if (StartCapCombo != null) StartCapCombo.SelectedIndex = 0;
            if (PathTypeCombo != null) PathTypeCombo.SelectedIndex = 1; // 默认为 Solid
            if (EndCapCombo != null) EndCapCombo.SelectedIndex = 0;
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
        // ✨ 完善后的预览方法：支持虚线比例实时预览
        private void UpdatePreview()
        {
            if (PreviewCanvas == null || _currentStyle == null) return;
            PreviewCanvas.Children.Clear();

            var brush = new SolidColorBrush(_currentStyle.MainColor);
            double centerY = PreviewCanvas.Height / 2;
            double startX = 60;
            double endX = 240;
            double m = 4.0; // 预览视觉缩放倍数

            if (_currentStyle.PathType == PathCategory.Dashed)
            {
                // 1. 虚线预览：动态计算虚线数组
                // 将 LtScaleSlider 的值映射到 WPF 的 DashArray
                // 基准值设为 3和2，随 LinetypeScale 线性缩放
                double dashValue = 3 * _currentStyle.LinetypeScale;
                double gapValue = 2 * _currentStyle.LinetypeScale;

                Path dashedPath = new Path
                {
                    Stroke = brush,
                    StrokeThickness = (_currentStyle.StartWidth + _currentStyle.EndWidth) / 2 * m,
                    // ✨ 关键点：动态绑定比例
                    StrokeDashArray = new DoubleCollection { dashValue, gapValue },
                    Data = new LineGeometry(new Point(startX, centerY), new Point(endX, centerY)),
                    Opacity = 0.8
                };
                PreviewCanvas.Children.Add(dashedPath);
            }
            else if (_currentStyle.PathType == PathCategory.Solid)
            {
                // 2. 实线预览：贝塞尔宽度模拟
                Polygon body = new Polygon { Fill = brush, Opacity = 0.8 };
                double wStart = _currentStyle.StartWidth * m;
                double wMid = _currentStyle.MidWidth * m;
                double wEnd = _currentStyle.EndWidth * m;
                double midX = (startX + endX) / 2;

                body.Points.Add(new Point(startX, centerY - wStart));
                body.Points.Add(new Point(midX, centerY - wMid));
                body.Points.Add(new Point(endX, centerY - wEnd));
                body.Points.Add(new Point(endX, centerY + wEnd));
                body.Points.Add(new Point(midX, centerY + wMid));
                body.Points.Add(new Point(startX, centerY + wStart));
                PreviewCanvas.Children.Add(body);
            }

            // 绘制端头逻辑
            double s = _currentStyle.ArrowSize;
            if (_currentStyle.EndCapStyle != ArrowHeadType.None)
            {
                Polygon head = new Polygon { Fill = brush };
                head.Points.Add(new Point(endX + s, centerY));
                head.Points.Add(new Point(endX, centerY - s * 0.4));
                head.Points.Add(new Point(endX, centerY + s * 0.4));
                PreviewCanvas.Children.Add(head);
            }

            if (_currentStyle.StartCapStyle == ArrowHeadType.Circle)
            {
                Ellipse dot = new Ellipse { Fill = brush, Width = 8, Height = 8 };
                Canvas.SetLeft(dot, startX - 4);
                Canvas.SetTop(dot, centerY - 4);
                PreviewCanvas.Children.Add(dot);
            }
        }

        // ✨ 修复 NullReferenceException 并实现三段式参数同步
        // 请替换完整的 SyncStyleFromUI 方法
        private void SyncStyleFromUI()
        {
            if (_currentStyle == null) _currentStyle = new AnalysisStyle();

            // 严谨的空检查，去掉了已不存在的 SegmentStyleCombo
            if (SizeSlider == null || StartWidthSlider == null || EndWidthSlider == null ||
                TransSlider == null || StartCapCombo == null || EndCapCombo == null ||
                PathTypeCombo == null || MidWidthSlider == null || LtScaleSlider == null)
                return;

            // 1. 同步端头尺寸
            _currentStyle.ArrowSize = SizeSlider.Value;

            // 2. 同步物理宽度
            _currentStyle.StartWidth = StartWidthSlider.Value;
            _currentStyle.EndWidth = EndWidthSlider.Value;
            _currentStyle.MidWidth = MidWidthSlider.Value;

            // 3. 同步组合样式
            if (StartCapCombo.SelectedIndex != -1)
                _currentStyle.StartCapStyle = (ArrowHeadType)StartCapCombo.SelectedIndex;

            if (EndCapCombo.SelectedIndex != -1)
                _currentStyle.EndCapStyle = (ArrowHeadType)EndCapCombo.SelectedIndex;

            // ✨ 使用 PathTypeCombo 控制路径分类
            if (PathTypeCombo.SelectedIndex != -1)
                _currentStyle.PathType = (PathCategory)PathTypeCombo.SelectedIndex;

            // 4. 同步其他参数
            _currentStyle.LinetypeScale = LtScaleSlider.Value;
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