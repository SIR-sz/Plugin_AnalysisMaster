using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
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
        // ✨ 完整方法：在 CAD 窗口中更新实时草稿
        private DBObjectCollection _transientEntities = new DBObjectCollection();

        private void UpdateCADLivePreview()
        {
            var tm = Autodesk.AutoCAD.GraphicsInterface.TransientManager.CurrentTransientManager;

            // 1. 清除旧的瞬态图形
            if (_transientEntities.Count > 0)
            {
                tm.EraseTransients(Autodesk.AutoCAD.GraphicsInterface.TransientDrawingMode.Main, 128, new IntegerCollection());
                foreach (DBObject obj in _transientEntities) obj.Dispose();
                _transientEntities.Clear();
            }

            // 2. 获取当前视图的中心点作为预览位置（或者使用最后一次点击的点）
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            Point3d viewCenter = (Point3d)Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("VIEWCTR");
            Point3d startPt = new Point3d(viewCenter.X - 50, viewCenter.Y, 0);
            Point3d endPt = new Point3d(viewCenter.X + 50, viewCenter.Y, 0);

            // 3. 生成新图元并添加到瞬态显示
            _transientEntities = GeometryEngine.GeneratePreviewEntities(startPt, endPt, _currentStyle);

            foreach (Entity ent in _transientEntities)
            {
                tm.AddTransient(ent, Autodesk.AutoCAD.GraphicsInterface.TransientDrawingMode.Main, 128, new IntegerCollection());
            }
        }
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

            this.Loaded += (s, e) =>
            {
                SyncStyleFromUI();
                UpdatePreview();
            };

            // ✨ 修复：移除了不存在的 SegmentStyleCombo 引用
            if (StartCapCombo != null) StartCapCombo.SelectedIndex = 0;
            if (PathTypeCombo != null) PathTypeCombo.SelectedIndex = 1;
            if (EndCapCombo != null) EndCapCombo.SelectedIndex = 0;
        }

        // ✨ 核心方法：参数改变触发预览更新
        // ✨ 完整方法：参数改变触发双重预览同步
        // ✨ 完整方法：回归方案 A，仅更新窗口内预览并清理 CAD 残留
        private void OnParamChanged(object sender, System.Windows.RoutedEventArgs e)
        {
            if (!this.IsLoaded) return;

            // 1. 同步数据
            SyncStyleFromUI();

            // 2. 仅更新 WPF 预览
            UpdatePreview();

            // 3. 清理 CAD 中可能残留的瞬态预览图形（方案 B 的清理）
            var tm = Autodesk.AutoCAD.GraphicsInterface.TransientManager.CurrentTransientManager;
            if (_transientEntities != null && _transientEntities.Count > 0)
            {
                tm.EraseTransients(Autodesk.AutoCAD.GraphicsInterface.TransientDrawingMode.Main, 128, new Autodesk.AutoCAD.Geometry.IntegerCollection());
                foreach (Autodesk.AutoCAD.DatabaseServices.DBObject obj in _transientEntities) obj.Dispose();
                _transientEntities.Clear();
            }
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
        // ✨ 完整方法：重构预览逻辑，对齐 AutoCAD 的虚线比例感，实现所见即所得
        // ✨ 完整方法：重构预览逻辑，深度对齐 CAD 比例感，解决“数值 0.1 偏差”问题
        private void UpdatePreview()
        {
            if (PreviewCanvas == null || _currentStyle == null) return;
            PreviewCanvas.Children.Clear();

            var brush = new System.Windows.Media.SolidColorBrush(_currentStyle.MainColor);
            double centerY = PreviewCanvas.Height / 2;
            double startX = 60;
            double endX = 240;
            double m = 4.0; // 预览视觉缩放倍数

            // 计算平均物理宽度用于抵消 WPF 的虚线缩放特性
            double avgWidth = (_currentStyle.StartWidth + _currentStyle.EndWidth) / 2.0;
            if (avgWidth < 0.1) avgWidth = 0.1;

            if (_currentStyle.PathType == PathCategory.Dashed)
            {
                // ✨ 核心修正：
                // 既然 CAD 需要调到 0.1 才有效果，说明预览里的“基准间距”太小了。
                // 我们将预览的基准值放大 10 倍（从 12 提升到 120），
                // 这样当滑块在 1.0 时，预览里会是一条极长的实线（模拟 CAD 里的实线感）；
                // 只有当用户把滑块拉低到 0.1 时，预览里的虚线段才会缩小到 12 像素左右，看起来才像虚线。
                double baseDash = 120.0;
                double baseGap = 60.0;

                // 计算预览用的 DashArray
                // 通过除以宽度 avgWidth，确保“线变粗”时，虚线间距不会像 WPF 默认那样变长
                double dashValue = (baseDash * _currentStyle.LinetypeScale) / avgWidth;
                double gapValue = (baseGap * _currentStyle.LinetypeScale) / avgWidth;

                System.Windows.Shapes.Path dashedPath = new System.Windows.Shapes.Path
                {
                    Stroke = brush,
                    StrokeThickness = avgWidth * m,
                    // 应用修正后的虚线比例
                    StrokeDashArray = new System.Windows.Media.DoubleCollection { dashValue, gapValue },
                    Data = new System.Windows.Media.LineGeometry(new System.Windows.Point(startX, centerY), new System.Windows.Point(endX, centerY)),
                    Opacity = 0.8
                };
                PreviewCanvas.Children.Add(dashedPath);
            }
            else if (_currentStyle.PathType == PathCategory.Solid)
            {
                // 实线类（含束腰效果）：采用高精度采样模拟
                System.Windows.Shapes.Polygon body = new System.Windows.Shapes.Polygon { Fill = brush, Opacity = 0.8 };
                int samples = 30;
                double totalLen = endX - startX;

                for (int i = 0; i <= samples; i++)
                {
                    double t = (double)i / samples;
                    double curX = startX + totalLen * t;
                    double curW = CalculateBezierWidth(t, _currentStyle.StartWidth, _currentStyle.MidWidth, _currentStyle.EndWidth) * m;
                    body.Points.Add(new System.Windows.Point(curX, centerY - curW));
                }
                for (int i = samples; i >= 0; i--)
                {
                    double t = (double)i / samples;
                    double curX = startX + totalLen * t;
                    double curW = CalculateBezierWidth(t, _currentStyle.StartWidth, _currentStyle.MidWidth, _currentStyle.EndWidth) * m;
                    body.Points.Add(new System.Windows.Point(curX, centerY + curW));
                }
                PreviewCanvas.Children.Add(body);
            }

            // 绘制端头
            double s = _currentStyle.ArrowSize;
            if (_currentStyle.EndCapStyle != ArrowHeadType.None)
            {
                System.Windows.Shapes.Polygon head = new System.Windows.Shapes.Polygon { Fill = brush };
                head.Points.Add(new System.Windows.Point(endX + s, centerY));
                head.Points.Add(new System.Windows.Point(endX, centerY - s * 0.4));
                head.Points.Add(new System.Windows.Point(endX, centerY + s * 0.4));
                PreviewCanvas.Children.Add(head);
            }

            if (_currentStyle.StartCapStyle == ArrowHeadType.Circle)
            {
                System.Windows.Shapes.Ellipse dot = new System.Windows.Shapes.Ellipse { Fill = brush, Width = 8, Height = 8 };
                System.Windows.Controls.Canvas.SetLeft(dot, startX - 4);
                System.Windows.Controls.Canvas.SetTop(dot, centerY - 4);
                PreviewCanvas.Children.Add(dot);
            }
        }

        // ✨ 辅助算法：确保预览与 CAD 使用同一套数学模型
        private double CalculateBezierWidth(double t, double start, double mid, double end)
        {
            double invT = 1.0 - t;
            return (invT * invT * start) + (2 * t * invT * mid) + (t * t * end);
        }

        // ✨ 修复 NullReferenceException 并实现三段式参数同步
        // 请替换完整的 SyncStyleFromUI 方法
        private void SyncStyleFromUI()
        {
            if (_currentStyle == null) _currentStyle = new AnalysisStyle();

            // 严谨的空检查，包含所有新加入的滑块
            if (SizeSlider == null || StartWidthSlider == null || EndWidthSlider == null ||
                MidWidthSlider == null || LtScaleSlider == null || TransSlider == null ||
                StartCapCombo == null || EndCapCombo == null || PathTypeCombo == null)
                return;

            // 1. 同步几何尺寸
            _currentStyle.ArrowSize = SizeSlider.Value;
            _currentStyle.StartWidth = StartWidthSlider.Value;
            _currentStyle.MidWidth = MidWidthSlider.Value;
            _currentStyle.EndWidth = EndWidthSlider.Value;

            // 2. 同步样式枚举
            if (StartCapCombo.SelectedIndex != -1)
                _currentStyle.StartCapStyle = (ArrowHeadType)StartCapCombo.SelectedIndex;

            if (EndCapCombo.SelectedIndex != -1)
                _currentStyle.EndCapStyle = (ArrowHeadType)EndCapCombo.SelectedIndex;

            if (PathTypeCombo.SelectedIndex != -1)
                _currentStyle.PathType = (PathCategory)PathTypeCombo.SelectedIndex;

            // 3. 同步渲染参数
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

        // ✨ 完整方法：窗口关闭逻辑，彻底断开与 CAD 瞬态管理器的联系
        private void Close_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            // 强制清理瞬态图形
            var tm = Autodesk.AutoCAD.GraphicsInterface.TransientManager.CurrentTransientManager;
            if (_transientEntities != null && _transientEntities.Count > 0)
            {
                tm.EraseTransients(Autodesk.AutoCAD.GraphicsInterface.TransientDrawingMode.Main, 128, new Autodesk.AutoCAD.Geometry.IntegerCollection());
                foreach (Autodesk.AutoCAD.DatabaseServices.DBObject obj in _transientEntities) obj.Dispose();
                _transientEntities.Clear();
            }

            this.Close();
        }
        // ✨ 完整方法：调用 AutoCAD 原生线型选择对话框并记录结果
        private void SelectLinetype_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            // 发送命令打开官方线型管理器 (等同于 LISP 的 (command "_.LINETYPE"))
            // 用户在窗口里选好并点击“置为当前”即可
            doc.SendStringToExecute("_.LINETYPE ", true, false, false);

            // 提示用户
            doc.Editor.WriteMessage("\n[动线专家] 请在弹出窗口中选择线型并点击“置为当前”。");

            // 自动读取当前设置
            RefreshCurrentLinetype();
        }
        // 辅助方法：读取 CAD 当前正在使用的线型名称
        private void RefreshCurrentLinetype()
        {
            Database db = HostApplicationServices.WorkingDatabase;
            using (var tr = db.TransactionManager.StartTransaction())
            {
                // 读取系统变量 CELTYPE (当前线型)
                string currentLt = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("CELTYPE").ToString();

                if (_currentStyle != null)
                {
                    _currentStyle.SelectedLinetype = currentLt;
                    UpdatePreview();
                }
                tr.Commit();
            }
        }
        // 辅助方法：验证并应用线型
        private void ApplyLinetypeByName(Database db, string ltName, Autodesk.AutoCAD.EditorInput.Editor ed)
        {
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var lt = (LinetypeTable)tr.GetObject(db.LinetypeTableId, OpenMode.ForRead);

                if (lt.Has(ltName))
                {
                    // 更新模型数据
                    _currentStyle.SelectedLinetype = ltName;

                    // 自动切换 UI 到虚线类
                    if (PathTypeCombo != null) PathTypeCombo.SelectedIndex = 2;
                    _currentStyle.PathType = PathCategory.Dashed;

                    // 同步与预览刷新
                    SyncStyleFromUI();
                    UpdatePreview();

                    ed.WriteMessage($"\n[动线专家] 线型已更改为: {ltName}");
                }
                else
                {
                    ed.WriteMessage($"\n[警告] 图纸中不存在线型 '{ltName}'，请先使用 LINETYPE 命令加载。");
                }
                tr.Commit();
            }
        }
        // 当路径类型改变时，切换面板显示状态
        // ✨ 完整代码块：添加到 MainControlWindow 类中
        // ✨ 完整修复版：确保分类引用正确
        private void OnPathTypeChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!this.IsLoaded) return;
            SyncStyleFromUI();

            if (DashedOptionsPanel != null)
                DashedOptionsPanel.Visibility = (_currentStyle.PathType == PathCategory.Dashed)
                    ? System.Windows.Visibility.Visible
                    : System.Windows.Visibility.Collapsed;

            if (PatternOptionsPanel != null)
                PatternOptionsPanel.Visibility = (_currentStyle.PathType == PathCategory.Pattern)
                    ? System.Windows.Visibility.Visible
                    : System.Windows.Visibility.Collapsed;

            UpdatePreview();
        }

        // 拾取图纸中的块参照
        // ✨ 完整方法：实现 CAD 块拾取逻辑
        private void PickBlock_Click(object sender, RoutedEventArgs e)
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            this.Hide();
            try
            {
                var opt = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\n请选择图纸中的块参照: ");
                opt.SetRejectMessage("\n所选对象必须是块参照！");
                opt.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.BlockReference), true);

                var res = doc.Editor.GetEntity(opt);
                if (res.Status == Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                {
                    using (var tr = doc.Database.TransactionManager.StartTransaction())
                    {
                        var br = (Autodesk.AutoCAD.DatabaseServices.BlockReference)tr.GetObject(res.ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                        _currentStyle.CustomBlockName = br.Name;
                        if (BlockNameTxt != null) BlockNameTxt.Text = br.Name;
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


    }
}