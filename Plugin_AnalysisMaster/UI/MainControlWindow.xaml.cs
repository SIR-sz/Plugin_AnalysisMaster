using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Plugin_AnalysisMaster.Models;
using Plugin_AnalysisMaster.Services;
using System;
using System.IO;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms; // 用于 DialogResult
using System.Windows.Media;
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
            LoadPatternLibrary();
            this.Loaded += (s, e) => { SyncStyleFromUI(); UpdatePreview(); };
        }

        private void LoadPatternLibrary()
        {
            if (BlockLibraryCombo == null) return;
            BlockLibraryCombo.Items.Clear();
            string dllPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string libPath = Path.Combine(dllPath, "Assets", "PatternLibrary.dwg");

            if (File.Exists(libPath))
            {
                using (Database db = new Database(false, true))
                {
                    db.ReadDwgFile(libPath, FileOpenMode.OpenForReadAndReadShare, true, "");
                    using (var tr = db.TransactionManager.StartTransaction())
                    {
                        BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                        foreach (ObjectId id in bt)
                        {
                            BlockTableRecord btr = (BlockTableRecord)tr.GetObject(id, OpenMode.ForRead);
                            if (!btr.IsLayout && !btr.IsAnonymous) BlockLibraryCombo.Items.Add(btr.Name);
                        }
                    }
                }
            }
            if (BlockLibraryCombo.Items.Count > 0) BlockLibraryCombo.SelectedIndex = 0;
        }

        // ✨ 完整修复：同步所有滑块数值到样式模型
        private void SyncStyleFromUI()
        {
            if (PathTypeCombo == null || _currentStyle == null) return;

            // 同步路径模式
            var pathItem = (ComboBoxItem)PathTypeCombo.SelectedItem;
            if (pathItem != null)
                _currentStyle.PathType = (PathCategory)Enum.Parse(typeof(PathCategory), pathItem.Tag.ToString());

            // 同步端头样式
            if (EndCapCombo != null && EndCapCombo.SelectedItem != null)
            {
                string capText = (EndCapCombo.SelectedItem as ComboBoxItem)?.Content.ToString();
                _currentStyle.EndCapStyle = capText == "基础箭头" ? ArrowHeadType.Basic :
                                            capText == "圆点" ? ArrowHeadType.Circle : ArrowHeadType.None;
            }

            // 同步数值参数
            _currentStyle.StartWidth = StartWidthSlider?.Value ?? 1.0;
            _currentStyle.MidWidth = MidWidthSlider?.Value ?? 0.8;
            _currentStyle.EndWidth = EndWidthSlider?.Value ?? 0.5;
            _currentStyle.ArrowSize = ArrowSizeSlider?.Value ?? 8.0;

            // ✨ 同步选中的块名
            _currentStyle.SelectedBlockName = BlockLibraryCombo?.SelectedItem?.ToString() ?? "";

            _currentStyle.PatternSpacing = SpacingSlider?.Value ?? 10.0;
            _currentStyle.PatternScale = PatternScaleSlider?.Value ?? 1.0;
            _currentStyle.Transparency = TransSlider?.Value ?? 0;
        }

        private void OnPathTypeChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!this.IsLoaded) return;
            SyncStyleFromUI();

            // ✨ 修复静态引用：使用全限定名
            if (SolidPanel != null)
                SolidPanel.Visibility = (_currentStyle.PathType == PathCategory.Solid)
                    ? System.Windows.Visibility.Visible : System.Windows.Visibility.Collapsed;

            if (PatternPanel != null)
                PatternPanel.Visibility = (_currentStyle.PathType == PathCategory.Pattern)
                    ? System.Windows.Visibility.Visible : System.Windows.Visibility.Collapsed;

            UpdatePreview();
        }

        private void StartDraw_Click(object sender, RoutedEventArgs e)
        {
            SyncStyleFromUI();
            this.Hide();
            try { GeometryEngine.DrawAnalysisLine(null, _currentStyle); }
            finally { this.Show(); }
        }

        // ✨ 修复：确保在 AddAllowedClass 之前调用 SetRejectMessage 避免崩溃
        private void PickBlock_Click(object sender, RoutedEventArgs e)
        {
            var doc = AcApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            this.Hide();
            try
            {
                var opt = new PromptEntityOptions("\n请选择图中的块作为样式单元: ");
                opt.SetRejectMessage("\n所选对象不是块参照，请重新选择！");
                opt.AddAllowedClass(typeof(BlockReference), true);

                var res = doc.Editor.GetEntity(opt);
                if (res.Status == PromptStatus.OK)
                {
                    using (var tr = doc.Database.TransactionManager.StartTransaction())
                    {
                        var br = (BlockReference)tr.GetObject(res.ObjectId, OpenMode.ForRead);
                        _currentStyle.CustomBlockName = br.Name;
                        if (BlockLibraryCombo != null)
                        {
                            if (!BlockLibraryCombo.Items.Contains(br.Name))
                                BlockLibraryCombo.Items.Add(br.Name);
                            BlockLibraryCombo.SelectedItem = br.Name;
                        }
                        tr.Commit();
                    }
                }
            }
            finally
            {
                this.Show();
                UpdatePreview();
            }
        }

        // ✨ 实现：预览窗口实时绘图逻辑
        // ✨ 完整实现：实时刷新预览画布，支持实线渐变和阵列样式预览
        private void UpdatePreview()
        {
            if (PreviewCanvas == null || !this.IsLoaded) return;

            // 1. 清空当前画布内容
            PreviewCanvas.Children.Clear();

            double width = PreviewCanvas.ActualWidth;
            double height = PreviewCanvas.ActualHeight;
            if (width <= 0 || height <= 0) return;

            // 2. 定义一段示意性的 S 型贝塞尔路径点
            System.Windows.Point p1 = new System.Windows.Point(width * 0.1, height * 0.7);
            System.Windows.Point p2 = new System.Windows.Point(width * 0.5, height * 0.1);
            System.Windows.Point p3 = new System.Windows.Point(width * 0.9, height * 0.5);

            // 3. 准备画刷和透明度
            var brush = new SolidColorBrush(_currentStyle.MainColor);
            brush.Opacity = (100 - _currentStyle.Transparency) / 100.0;

            // 4. 根据绘制模式渲染主体
            if (_currentStyle.PathType == PathCategory.Solid)
            {
                // 渲染连续实线：通过分段绘制模拟宽度渐变
                int segments = 50;
                for (int i = 0; i < segments; i++)
                {
                    double t1 = i / (double)segments;
                    double t2 = (i + 1) / (double)segments;

                    System.Windows.Point pt1 = GetBezierPoint(t1, p1, p2, p3);
                    System.Windows.Point pt2 = GetBezierPoint(t2, p1, p2, p3);

                    // 获取该进度下的贝塞尔插值宽度
                    double w = CalculateBezierWidth(t1, _currentStyle.StartWidth, _currentStyle.MidWidth, _currentStyle.EndWidth);

                    var line = new System.Windows.Shapes.Line
                    {
                        X1 = pt1.X,
                        Y1 = pt1.Y,
                        X2 = pt2.X,
                        Y2 = pt2.Y,
                        Stroke = brush,
                        StrokeThickness = w * 1.5, // 预览稍微加粗方便观察
                        StrokeStartLineCap = PenLineCap.Round,
                        StrokeEndLineCap = PenLineCap.Round
                    };
                    PreviewCanvas.Children.Add(line);
                }
            }
            else if (_currentStyle.PathType == PathCategory.Pattern)
            {
                // 渲染阵列模式：根据选中的块名模拟不同形状
                string name = (_currentStyle.SelectedBlockName ?? "").ToLower();
                double step = Math.Max(0.05, 1.0 / (width / (_currentStyle.PatternSpacing + 1)));

                for (double t = 0; t <= 1; t += step)
                {
                    System.Windows.Point pt = GetBezierPoint(t, p1, p2, p3);
                    double s = _currentStyle.PatternScale * 6; // 缩放适配预览

                    FrameworkElement shape;
                    // 简单的关键词匹配，让预览感知“图块切换”
                    if (name.Contains("箭") || name.Contains("arrow"))
                    {
                        var poly = new System.Windows.Shapes.Polygon { Fill = brush };
                        poly.Points.Add(new System.Windows.Point(s, 0));
                        poly.Points.Add(new System.Windows.Point(-s, s / 2));
                        poly.Points.Add(new System.Windows.Point(-s, -s / 2));
                        shape = poly;
                    }
                    else if (name.Contains("线") || name.Contains("dash"))
                    {
                        shape = new System.Windows.Shapes.Rectangle { Width = s * 2, Height = s / 2, Fill = brush };
                    }
                    else
                    {
                        shape = new System.Windows.Shapes.Ellipse { Width = s, Height = s, Fill = brush };
                    }

                    Canvas.SetLeft(shape, pt.X - s / 2);
                    Canvas.SetTop(shape, pt.Y - s / 2);
                    PreviewCanvas.Children.Add(shape);
                }
            }

            // 5. 渲染端头箭头预览
            if (_currentStyle.EndCapStyle != ArrowHeadType.None)
            {
                System.Windows.Point endPt = GetBezierPoint(1.0, p1, p2, p3);
                System.Windows.Point preEndPt = GetBezierPoint(0.95, p1, p2, p3);

                System.Windows.Vector dir = endPt - preEndPt;
                if (dir.Length > 0) dir.Normalize();
                System.Windows.Vector norm = new System.Windows.Vector(-dir.Y, dir.X);

                double arrowSize = _currentStyle.ArrowSize * 1.2;
                var arrow = new System.Windows.Shapes.Polygon { Fill = brush };
                arrow.Points.Add(endPt);
                arrow.Points.Add(endPt - dir * arrowSize + norm * arrowSize * 0.4);
                arrow.Points.Add(endPt - dir * arrowSize - norm * arrowSize * 0.4);

                PreviewCanvas.Children.Add(arrow);
            }
        }

        // ✨ 辅助方法：计算二次贝塞尔曲线上的点
        private System.Windows.Point GetBezierPoint(double t, System.Windows.Point p0, System.Windows.Point p1, System.Windows.Point p2)
        {
            double x = (1 - t) * (1 - t) * p0.X + 2 * t * (1 - t) * p1.X + t * t * p2.X;
            double y = (1 - t) * (1 - t) * p0.Y + 2 * t * (1 - t) * p1.Y + t * t * p2.Y;
            return new System.Windows.Point(x, y);
        }
        private double CalculateBezierWidth(double t, double s, double m, double e)
        {
            return (1 - t) * (1 - t) * s + 2 * t * (1 - t) * m + t * t * e;
        }
        // ✨ 修复：确保所有参数变动时先同步数据模型

        private void OnParamChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (this.IsLoaded)
            {
                SyncStyleFromUI(); // 必须先同步数据
                UpdatePreview();
            }
        }
        // ✨ 修复：下拉列表切换时的处理 (解决您提到的不刷新问题)
        private void OnParamChanged(object sender, SelectionChangedEventArgs e)
        {
            if (this.IsLoaded)
            {
                SyncStyleFromUI(); // 必须先同步数据
                UpdatePreview();
            }
        }
        // ✨ 实现：点击颜色预览块，弹出 AutoCAD 标准颜色选择对话框
        private void SelectColor_Click(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            // 1. 初始化 AutoCAD 颜色对话框
            Autodesk.AutoCAD.Windows.ColorDialog colorDlg = new Autodesk.AutoCAD.Windows.ColorDialog();

            // 2. 设置对话框的初始颜色
            colorDlg.Color = Autodesk.AutoCAD.Colors.Color.FromRgb(
                _currentStyle.MainColor.R,
                _currentStyle.MainColor.G,
                _currentStyle.MainColor.B);

            // 3. 显示对话框
            if (colorDlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                // 4. 获取选中的颜色 (System.Drawing.Color)
                System.Drawing.Color selectedColor = colorDlg.Color.ColorValue;

                // 5. 更新数据模型
                _currentStyle.MainColor = System.Windows.Media.Color.FromRgb(
                    selectedColor.R,
                    selectedColor.G,
                    selectedColor.B);

                // 6. 更新 UI 预览块的颜色
                if (ColorPreview != null)
                {
                    ColorPreview.Fill = new SolidColorBrush(_currentStyle.MainColor);
                }

                // 7. 实时更新预览 Canvas
                UpdatePreview();
            }
        }
        private void Close_Click(object sender, RoutedEventArgs e) => this.Close();


    }
}