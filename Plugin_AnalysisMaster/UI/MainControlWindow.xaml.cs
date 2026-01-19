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
using System.Windows.Shapes;
using AcApp = Autodesk.AutoCAD.ApplicationServices.Application;
using Ellipse = System.Windows.Shapes.Ellipse; // 默认将 Ellipse 指向 WPF 形状
using Path = System.IO.Path; // 默认将 Path 指向文件路径工具

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
            string dllPath = System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string libPath = System.IO.Path.Combine(dllPath, "Assets", "PatternLibrary.dwg");

            if (System.IO.File.Exists(libPath))
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

            var item = (ComboBoxItem)PathTypeCombo.SelectedItem;
            if (item != null)
                _currentStyle.PathType = (PathCategory)Enum.Parse(typeof(PathCategory), item.Tag.ToString());

            // 同步所有宽度和尺寸
            _currentStyle.StartWidth = StartWidthSlider?.Value ?? 1.0;
            _currentStyle.MidWidth = MidWidthSlider?.Value ?? 0.8; // ✨ 补全
            _currentStyle.EndWidth = EndWidthSlider?.Value ?? 0.5;
            _currentStyle.ArrowSize = ArrowSizeSlider?.Value ?? 8.0;

            // 同步阵列参数
            _currentStyle.SelectedBlockName = BlockLibraryCombo?.SelectedItem?.ToString() ?? "";
            _currentStyle.PatternSpacing = SpacingSlider?.Value ?? 10.0;
            _currentStyle.PatternScale = PatternScaleSlider?.Value ?? 1.0; // ✨ 补全

            // 同步透明度
            _currentStyle.Transparency = TransSlider?.Value ?? 0; // ✨ 补全
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

        private void UpdatePreview()
        {
            if (PreviewCanvas == null || !this.IsLoaded) return;
            PreviewCanvas.Children.Clear();

            double w = PreviewCanvas.ActualWidth;
            double h = PreviewCanvas.ActualHeight;
            if (w <= 0 || h <= 0) return;

            // 1. 定义示意路径
            System.Windows.Point p1 = new System.Windows.Point(w * 0.1, h * 0.7);
            System.Windows.Point p2 = new System.Windows.Point(w * 0.5, h * 0.1);
            System.Windows.Point p3 = new System.Windows.Point(w * 0.9, h * 0.5);

            var brush = new SolidColorBrush(_currentStyle.MainColor);
            brush.Opacity = (100 - _currentStyle.Transparency) / 100.0;

            // 2. 绘制主体
            if (_currentStyle.PathType == PathCategory.Solid)
            {
                int segments = 40;
                for (int i = 0; i < segments; i++)
                {
                    double t1 = i / (double)segments;
                    double t2 = (i + 1) / (double)segments;
                    System.Windows.Point pt1 = GetBezierPoint(t1, p1, p2, p3);
                    System.Windows.Point pt2 = GetBezierPoint(t2, p1, p2, p3);
                    double thickness = CalculateBezierWidth(t1, _currentStyle.StartWidth, _currentStyle.MidWidth, _currentStyle.EndWidth);

                    var line = new System.Windows.Shapes.Line
                    {
                        X1 = pt1.X,
                        Y1 = pt1.Y,
                        X2 = pt2.X,
                        Y2 = pt2.Y,
                        Stroke = brush,
                        StrokeThickness = thickness * 1.5,
                        StrokeStartLineCap = PenLineCap.Round,
                        StrokeEndLineCap = PenLineCap.Round
                    };
                    PreviewCanvas.Children.Add(line);
                }
            }
            else if (_currentStyle.PathType == PathCategory.Pattern)
            {
                string name = (_currentStyle.SelectedBlockName ?? "").ToLower();
                for (double t = 0; t <= 1; t += 0.1)
                {
                    System.Windows.Point pt = GetBezierPoint(t, p1, p2, p3);
                    double s = _currentStyle.PatternScale * 5;
                    FrameworkElement shape;

                    if (name.Contains("箭") || name.Contains("arrow"))
                    {
                        var poly = new System.Windows.Shapes.Polygon { Fill = brush }; // 明确指定 Shapes
                        poly.Points.Add(new System.Windows.Point(s, 0));
                        poly.Points.Add(new System.Windows.Point(-s, s / 2));
                        poly.Points.Add(new System.Windows.Point(-s, -s / 2));
                        shape = poly;
                    }
                    else
                    {
                        // ✨ 修复：明确指定使用 System.Windows.Shapes.Ellipse
                        shape = new System.Windows.Shapes.Ellipse { Width = s, Height = s, Fill = brush };
                    }

                    Canvas.SetLeft(shape, pt.X - s / 2);
                    Canvas.SetTop(shape, pt.Y - s / 2);
                    PreviewCanvas.Children.Add(shape);
                }
            }
        }

        // ✨ 辅助方法 (确保整个类中只保留一份)
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
                SyncStyleFromUI(); // ✨ 必须先同步 UI 数据到 _currentStyle
                UpdatePreview();
            }
        }
        // ✨ 修复：下拉列表切换时的处理 (解决您提到的不刷新问题)
        private void OnParamChanged(object sender, SelectionChangedEventArgs e)
        {
            if (this.IsLoaded)
            {
                SyncStyleFromUI(); // ✨ 必须先同步 UI 数据到 _currentStyle
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