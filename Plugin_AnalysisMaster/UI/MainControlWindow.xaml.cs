using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Plugin_AnalysisMaster.Models;
using Plugin_AnalysisMaster.Services;
using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using AcApp = Autodesk.AutoCAD.ApplicationServices.Application;

// 别名定义，解决命名空间冲突
using Image = System.Windows.Controls.Image;
using Line = System.Windows.Shapes.Line;
using Path = System.IO.Path;

namespace Plugin_AnalysisMaster.UI
{
    public partial class MainControlWindow : Window
    {
        // 将 _currentStyle 改为 static 静态变量，使其生命周期与 CAD 进程同步，实现会话级持久化。
        private static AnalysisStyle _currentStyle = new AnalysisStyle();
        private static MainControlWindow _instance;

        [DllImport("gdi32.dll")]
        public static extern bool DeleteObject(IntPtr hObject);

        public static void ShowTool()
        {
            if (_instance == null || !_instance.IsLoaded)
            {
                _instance = new MainControlWindow();
                AcApp.ShowModelessWindow(_instance);
            }
            else { _instance.Activate(); }
        }

        /// <summary>
        /// 窗口构造函数。
        /// 修改逻辑：在窗口加载完成（Loaded）事件中增加 ApplyStyleToUI 调用，
        /// 确保窗口每次打开时都能从静态变量中恢复上一次的参数设置。
        /// </summary>
        public MainControlWindow()
        {
            InitializeComponent();

            // 1. 先加载资源库（填充下拉框 Items）
            LoadPatternLibrary();

            // 2. 注册加载事件：回填数据并刷新预览
            this.Loaded += (s, e) =>
            {
                // 从内存回填到 UI
                ApplyStyleToUI();
                // 刷新一次预览图
                UpdatePreview();
            };
        }
        /// <summary>
        /// 将静态模型中的数据回填到 UI 界面。
        /// 修改逻辑：增加了对 SkeletonTypeCombo 的状态恢复，确保重新打开面板时保留骨架类型设置。
        /// </summary>
        private void ApplyStyleToUI()
        {
            if (_currentStyle == null) return;

            // 1. 恢复渲染模式
            foreach (ComboBoxItem item in PathTypeCombo.Items)
            {
                if (item.Tag.ToString() == _currentStyle.PathType.ToString())
                {
                    PathTypeCombo.SelectedItem = item;
                    break;
                }
            }

            // 2. ✨ 恢复骨架类型选中项
            if (SkeletonTypeCombo != null)
            {
                foreach (ComboBoxItem item in SkeletonTypeCombo.Items)
                {
                    if (item.Tag.ToString().Equals(_currentStyle.IsCurved.ToString(), StringComparison.OrdinalIgnoreCase))
                    {
                        SkeletonTypeCombo.SelectedItem = item;
                        break;
                    }
                }
            }

            // 3. 恢复下拉框选中项 (阵列单元、起终点)
            if (BlockLibraryCombo != null)
            {
                foreach (ComboBoxItem item in BlockLibraryCombo.Items)
                    if (item.Tag.ToString() == _currentStyle.SelectedBlockName) { BlockLibraryCombo.SelectedItem = item; break; }
            }
            if (StartArrowCombo != null)
            {
                foreach (ComboBoxItem item in StartArrowCombo.Items)
                    if (item.Tag.ToString() == _currentStyle.StartArrowType) { StartArrowCombo.SelectedItem = item; break; }
            }
            if (EndArrowCombo != null)
            {
                foreach (ComboBoxItem item in EndArrowCombo.Items)
                    if (item.Tag.ToString() == _currentStyle.EndArrowType) { EndArrowCombo.SelectedItem = item; break; }
            }

            // 4. 恢复滑块与预览块 (保持不变...)
            if (StartWidthSlider != null) StartWidthSlider.Value = _currentStyle.StartWidth;
            if (MidWidthSlider != null) MidWidthSlider.Value = _currentStyle.MidWidth;
            if (EndWidthSlider != null) EndWidthSlider.Value = _currentStyle.EndWidth;
            if (SpacingSlider != null) SpacingSlider.Value = _currentStyle.PatternSpacing;
            if (PatternScaleSlider != null) PatternScaleSlider.Value = _currentStyle.PatternScale;
            if (ArrowSizeSlider != null) ArrowSizeSlider.Value = _currentStyle.ArrowSize;
            if (CapIndentSlider != null) CapIndentSlider.Value = _currentStyle.CapIndent;
            if (ColorPreview != null) ColorPreview.Fill = new SolidColorBrush(_currentStyle.MainColor);
        }
        /// <summary>
        /// 加载资源库并同时填充两个图元下拉框。
        /// 修改逻辑：增加了对 BlockLibraryCombo2 的初始化和填充。
        /// </summary>
        private void LoadPatternLibrary()
        {
            if (BlockLibraryCombo == null || BlockLibraryCombo2 == null || StartArrowCombo == null || EndArrowCombo == null) return;

            BlockLibraryCombo.Items.Clear();
            BlockLibraryCombo2.Items.Clear(); // 清理第二个下拉框
            StartArrowCombo.Items.Clear();
            EndArrowCombo.Items.Clear();

            StartArrowCombo.Items.Add(new ComboBoxItem { Content = "无", Tag = "None" });
            EndArrowCombo.Items.Add(new ComboBoxItem { Content = "无", Tag = "None" });
            StartArrowCombo.SelectedIndex = 0;
            EndArrowCombo.SelectedIndex = 0;

            string dllPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
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
                            if (btr.IsLayout || btr.IsAnonymous) continue;

                            string realName = btr.Name;
                            if (realName.StartsWith("Cap_", StringComparison.OrdinalIgnoreCase))
                            {
                                string displayName = realName.Substring(4);
                                StartArrowCombo.Items.Add(new ComboBoxItem { Content = displayName, Tag = realName });
                                EndArrowCombo.Items.Add(new ComboBoxItem { Content = displayName, Tag = realName });
                            }
                            else if (realName.StartsWith("Pat_", StringComparison.OrdinalIgnoreCase))
                            {
                                string displayName = realName.Substring(4);
                                BlockLibraryCombo.Items.Add(new ComboBoxItem { Content = displayName, Tag = realName });
                                BlockLibraryCombo2.Items.Add(new ComboBoxItem { Content = displayName, Tag = realName }); // 同步填充第二个
                            }
                        }
                    }
                }
            }
            if (BlockLibraryCombo.Items.Count > 0) BlockLibraryCombo.SelectedIndex = 0;
            if (BlockLibraryCombo2.Items.Count > 1) BlockLibraryCombo2.SelectedIndex = 1; // 默认选第二个不一样的
            else if (BlockLibraryCombo2.Items.Count > 0) BlockLibraryCombo2.SelectedIndex = 0;
        }

        /// <summary>
        /// 组合模式切换按钮点击事件。
        /// 修复了 CS0176 错误：显式指定 System.Windows.Visibility 枚举。
        /// </summary>
        private void OnCompositeToggle_Click(object sender, RoutedEventArgs e)
        {
            bool isComposite = CompositeCheckBox.IsChecked ?? false;

            // ✨ 修改：显式使用 System.Windows.Visibility 解决命名冲突
            CompositeSettingsPanel.Visibility = isComposite
                ? System.Windows.Visibility.Visible
                : System.Windows.Visibility.Collapsed;

            // 调用合并后的参数变更方法
            OnParamChanged(sender, e);
        }

        /// <summary>
        /// 将 UI 数值同步到静态模型。
        /// 修改逻辑：增加了对 IsComposite 和 SelectedBlockName2 的同步。
        /// </summary>
        private void SyncStyleFromUI()
        {
            if (PathTypeCombo == null || _currentStyle == null) return;

            if (PathTypeCombo.SelectedItem is ComboBoxItem typeItem)
                _currentStyle.PathType = (PathCategory)Enum.Parse(typeof(PathCategory), typeItem.Tag.ToString());

            if (SkeletonTypeCombo != null && SkeletonTypeCombo.SelectedItem is ComboBoxItem skelItem)
                _currentStyle.IsCurved = bool.Parse(skelItem.Tag.ToString());

            // 同步图元 1
            if (BlockLibraryCombo != null && BlockLibraryCombo.SelectedItem is ComboBoxItem patItem)
                _currentStyle.SelectedBlockName = patItem.Tag.ToString();

            // ✨ 同步组合模式及图元 2
            _currentStyle.IsComposite = CompositeCheckBox?.IsChecked ?? false;
            if (BlockLibraryCombo2 != null && BlockLibraryCombo2.SelectedItem is ComboBoxItem patItem2)
                _currentStyle.SelectedBlockName2 = patItem2.Tag.ToString();

            if (StartArrowCombo != null && StartArrowCombo.SelectedItem is ComboBoxItem startItem)
                _currentStyle.StartArrowType = startItem.Tag.ToString();

            if (EndArrowCombo != null && EndArrowCombo.SelectedItem is ComboBoxItem endItem)
                _currentStyle.EndArrowType = endItem.Tag.ToString();

            _currentStyle.StartWidth = StartWidthSlider?.Value ?? 1.0;
            _currentStyle.MidWidth = MidWidthSlider?.Value ?? 0.8;
            _currentStyle.EndWidth = EndWidthSlider?.Value ?? 1.0;
            _currentStyle.PatternSpacing = SpacingSlider?.Value ?? 2.0;
            _currentStyle.PatternScale = PatternScaleSlider?.Value ?? 1.0;
            _currentStyle.ArrowSize = ArrowSizeSlider?.Value ?? 1.0;
            _currentStyle.CapIndent = CapIndentSlider?.Value ?? 0.0;
        }
        /// <summary>
        /// 刷新预览画布的核心逻辑。
        /// 包含：
        /// 1. 连续线模式 (Solid)：支持实时宽度、缩进、骨架（曲线/折线）及视觉缩放预览。
        /// 2. 阵列样式模式 (Pattern)：静态示意模式，固定显示 5 个单元，支持“组合模式”下的 A-B-A-B-A 交替显示。
        /// 3. 自动计算视觉缩放系数，确保大参数下预览不溢出。
        /// </summary>
        private void UpdatePreview()
        {
            if (PreviewCanvas == null || !this.IsLoaded) return;
            PreviewCanvas.Children.Clear();

            double w = PreviewCanvas.ActualWidth;
            double h = PreviewCanvas.ActualHeight;
            if (w <= 0 || h <= 0) return;

            // 设定预览控制点范围
            System.Windows.Point p1 = new System.Windows.Point(w * 0.2, h * 0.7);
            System.Windows.Point p2 = new System.Windows.Point(w * 0.5, h * 0.2);
            System.Windows.Point p3 = new System.Windows.Point(w * 0.8, h * 0.5);

            var brush = new SolidColorBrush(_currentStyle.MainColor);

            // 定义路径获取逻辑（样条曲线 vs 折线）
            Func<double, System.Windows.Point> getPathPoint = (t) =>
            {
                if (_currentStyle.IsCurved) return GetBezierPoint(t, p1, p2, p3);
                if (t < 0.5) return new System.Windows.Point(p1.X + (p2.X - p1.X) * (t * 2), p1.Y + (p2.Y - p1.Y) * (t * 2));
                return new System.Windows.Point(p2.X + (p3.X - p2.X) * ((t - 0.5) * 2), p2.Y + (p3.Y - p2.Y) * ((t - 0.5) * 2));
            };

            // 分支处理：连续线模式 (Solid)
            if (_currentStyle.PathType == PathCategory.Solid)
            {
                // 计算视觉缩放（基于最大宽度限制在 40px 内）
                double maxW = Math.Max(_currentStyle.StartWidth, Math.Max(_currentStyle.MidWidth, _currentStyle.EndWidth));
                double visualScale = maxW > 40.0 ? 40.0 / maxW : 1.0;

                // 计算缩进效果
                double totalLenGuess = 200.0;
                double indentOffset = (_currentStyle.CapIndent * visualScale / totalLenGuess) * 0.5;
                indentOffset = Math.Max(-0.2, Math.Min(0.4, indentOffset));

                int segments = 40;
                double tStart = Math.Max(0, indentOffset);
                double tEnd = Math.Min(1, 1 - indentOffset);

                for (int i = 0; i < segments; i++)
                {
                    double t1 = tStart + (tEnd - tStart) * (i / (double)segments);
                    double t2 = tStart + (tEnd - tStart) * ((i + 1.0) / segments);
                    System.Windows.Point pt1 = getPathPoint(t1);
                    System.Windows.Point pt2 = getPathPoint(t2);
                    double thickness = CalculateBezierWidth(t1, _currentStyle.StartWidth, _currentStyle.MidWidth, _currentStyle.EndWidth) * visualScale;

                    Line line = new Line
                    {
                        X1 = pt1.X,
                        Y1 = pt1.Y,
                        X2 = pt2.X,
                        Y2 = pt2.Y,
                        Stroke = brush,
                        StrokeThickness = thickness,
                        StrokeStartLineCap = PenLineCap.Round,
                        StrokeEndLineCap = PenLineCap.Round
                    };
                    PreviewCanvas.Children.Add(line);
                }

                // 渲染实时箭头预览
                RenderArrowPreview(p1, 0.0, p1, p2, p3, _currentStyle.StartArrowType, brush, true, visualScale);
                RenderArrowPreview(p3, 1.0, p1, p2, p3, _currentStyle.EndArrowType, brush, false, visualScale);
            }
            // 分支处理：阵列样式模式 (Pattern)
            else if (_currentStyle.PathType == PathCategory.Pattern)
            {
                // 1. 渲染固定的 5 个中间单元
                if (!string.IsNullOrEmpty(_currentStyle.SelectedBlockName))
                {
                    // 同步主图元
                    SyncBlockToCurrentDoc(_currentStyle.SelectedBlockName);
                    ImageSource mask1 = GetBlockMaskSource(_currentStyle.SelectedBlockName);

                    // 同步组合图元
                    ImageSource mask2 = null;
                    if (_currentStyle.IsComposite && !string.IsNullOrEmpty(_currentStyle.SelectedBlockName2))
                    {
                        SyncBlockToCurrentDoc(_currentStyle.SelectedBlockName2);
                        mask2 = GetBlockMaskSource(_currentStyle.SelectedBlockName2);
                    }

                    double fixedBlockSize = 20.0;
                    // 均匀分布的 5 个点位 (t = 1/6, 2/6, 3/6, 4/6, 5/6)
                    double[] fixedT = { 0.167, 0.333, 0.5, 0.667, 0.833 };

                    for (int i = 0; i < fixedT.Length; i++)
                    {
                        // 如果开启组合模式，奇数索引位显示图元 2
                        ImageSource currentMask = (i % 2 != 0 && _currentStyle.IsComposite && mask2 != null) ? mask2 : mask1;

                        System.Windows.Point pt = getPathPoint(fixedT[i]);
                        RenderPreviewItem(pt, fixedT[i], p1, p2, p3, currentMask, brush, fixedBlockSize);
                    }
                }

                // 2. 渲染固定的起终点箭头 (visualScale = -1 表示强制固定大小模式)
                RenderArrowPreview(p1, 0.0, p1, p2, p3, _currentStyle.StartArrowType, brush, true, -1.0);
                RenderArrowPreview(p3, 1.0, p1, p2, p3, _currentStyle.EndArrowType, brush, false, -1.0);
            }
        }
        /// <summary>
        /// “管理资源库”按钮点击事件。
        /// 作用：获取插件所在目录下的 Assets 文件夹路径，并调用系统资源管理器直接打开，
        /// 方便用户快速找到并编辑 PatternLibrary.dwg 文件。
        /// </summary>
        private void OpenLibraryFolder_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 1. 获取当前 DLL 运行所在的路径
                string dllPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);

                // 2. 定位到存放 DWG 库的 Assets 文件夹
                string assetsPath = System.IO.Path.Combine(dllPath, "Assets");

                if (System.IO.Directory.Exists(assetsPath))
                {
                    // 3. 调用系统进程打开资源管理器并定位到该目录
                    System.Diagnostics.Process.Start("explorer.exe", assetsPath);
                }
                else
                {
                    Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog("未找到 Assets 文件夹，请检查插件安装是否完整。");
                }
            }
            catch (Exception ex)
            {
                Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog("无法打开文件夹：" + ex.Message);
            }
        }
        /// <summary>
        /// 渲染预览图中的路径单元图块。
        /// 修改说明：保持角度计算逻辑，确保无论在实时模式还是固定模式下，
        /// 图块都能正确跟随骨架线（曲线或折线）的方向进行旋转。
        /// </summary>
        private void RenderPreviewItem(System.Windows.Point pt, double t, System.Windows.Point p0, System.Windows.Point p1, System.Windows.Point p2, ImageSource mask, Brush brush, double size)
        {
            if (mask == null) return;

            // 安全下限
            if (size < 2) size = 2;

            var rect = new System.Windows.Shapes.Rectangle { Width = size, Height = size, Fill = brush, OpacityMask = new ImageBrush(mask) };

            // ✨ 动态切线旋转逻辑
            Vector v;
            if (_currentStyle.IsCurved)
            {
                System.Windows.Point nPt = GetBezierPoint(Math.Min(t + 0.01, 1), p0, p1, p2);
                v = nPt - pt;
            }
            else
            {
                // 折线模式方向判断
                v = (t < 0.5) ? (p1 - p0) : (p2 - p1);
            }

            double angle = Math.Atan2(v.Y, v.X) * 180 / Math.PI;

            rect.RenderTransform = new RotateTransform(angle, size / 2, size / 2);
            Canvas.SetLeft(rect, pt.X - size / 2);
            Canvas.SetTop(rect, pt.Y - size / 2);
            PreviewCanvas.Children.Add(rect);
        }

        /// <summary>
        /// 渲染起终点箭头的预览效果。
        /// 修改说明：增加对固定大小模式的支持。如果传入的 visualScale < 0，则忽略全局缩放，
        /// 使用预设的固定尺寸（24像素），确保在阵列模式下箭头清晰且不被遮挡。
        /// </summary>
        private void RenderArrowPreview(System.Windows.Point pt, double t, System.Windows.Point p0, System.Windows.Point p1, System.Windows.Point p2, string blockName, Brush brush, bool isStart, double visualScale)
        {
            if (string.IsNullOrEmpty(blockName) || blockName == "None") return;

            SyncBlockToCurrentDoc(blockName);
            ImageSource mask = GetBlockMaskSource(blockName);
            if (mask == null) return;

            // ✨ 核心逻辑：判断是否为固定大小模式（用于阵列示意）
            double size;
            if (visualScale < 0)
            {
                size = 24.0; // 阵列模式下固定 24 像素
            }
            else
            {
                // 连续线模式下根据 ArrowSize 实时缩放
                size = 24 * (_currentStyle.ArrowSize / 8.0) * visualScale;
                if (size < 4) size = 4;
            }

            var rect = new System.Windows.Shapes.Rectangle { Width = size, Height = size, Fill = brush, OpacityMask = new ImageBrush(mask) };

            // 计算切线角度
            Vector v;
            if (_currentStyle.IsCurved)
            {
                System.Windows.Point nextPt = GetBezierPoint(isStart ? 0.01 : 1.0, p0, p1, p2);
                System.Windows.Point prevPt = GetBezierPoint(isStart ? 0.0 : 0.99, p0, p1, p2);
                v = nextPt - prevPt;
            }
            else
            {
                v = isStart ? (p1 - p0) : (p2 - p1);
            }

            double angle = Math.Atan2(v.Y, v.X) * 180 / Math.PI;
            if (isStart) angle += 180;

            rect.RenderTransform = new RotateTransform(angle, size / 2, size / 2);
            Canvas.SetLeft(rect, pt.X - size / 2);
            Canvas.SetTop(rect, pt.Y - size / 2);
            PreviewCanvas.Children.Add(rect);
        }
        private void SyncBlockToCurrentDoc(string blockName)
        {
            var doc = AcApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            // ✨ 修复：锁定文档，防止 eLockViolation 错误
            using (DocumentLock loc = doc.LockDocument())
            {
                using (var tr = doc.Database.TransactionManager.StartTransaction())
                {
                    BlockTable bt = (BlockTable)tr.GetObject(doc.Database.BlockTableId, OpenMode.ForRead);
                    if (bt.Has(blockName)) return;

                    string dllPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                    string libPath = Path.Combine(dllPath, "Assets", "PatternLibrary.dwg");

                    if (File.Exists(libPath))
                    {
                        using (Database srcDb = new Database(false, true))
                        {
                            srcDb.ReadDwgFile(libPath, FileShare.Read, true, "");
                            ObjectIdCollection ids = new ObjectIdCollection();
                            using (var trSrc = srcDb.TransactionManager.StartTransaction())
                            {
                                BlockTable btSrc = (BlockTable)trSrc.GetObject(srcDb.BlockTableId, OpenMode.ForRead);
                                if (btSrc.Has(blockName)) ids.Add(btSrc[blockName]);
                                trSrc.Commit();
                            }
                            if (ids.Count > 0)
                            {
                                IdMapping mapping = new IdMapping();
                                doc.Database.WblockCloneObjects(ids, doc.Database.BlockTableId, mapping, DuplicateRecordCloning.Ignore, false);
                            }
                        }
                    }
                    tr.Commit();
                }
            }
        }

        // ✨ 改进：获取带透明背景的块快照作为遮罩
        private ImageSource GetBlockMaskSource(string blockName)
        {
            var doc = AcApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return null;

            using (var tr = doc.Database.TransactionManager.StartTransaction())
            {
                BlockTable bt = (BlockTable)tr.GetObject(doc.Database.BlockTableId, OpenMode.ForRead);
                if (bt.Has(blockName))
                {
                    var acColor = Autodesk.AutoCAD.Colors.Color.FromRgb(0, 0, 0); // 黑色背景
                    IntPtr hBitmap = Autodesk.AutoCAD.Internal.Utils.GetBlockImage(bt[blockName], 128, 128, acColor);

                    if (hBitmap != IntPtr.Zero)
                    {
                        try
                        {
                            // 将 HBitmap 转为 System.Drawing.Bitmap 以便处理透明度
                            using (System.Drawing.Bitmap bmp = System.Drawing.Image.FromHbitmap(hBitmap))
                            {
                                // ✨ 关键：将黑色背景设为透明，否则预览就是方形色块
                                bmp.MakeTransparent(System.Drawing.Color.Black);
                                return BitmapToImageSource(bmp);
                            }
                        }
                        finally { DeleteObject(hBitmap); }
                    }
                }
            }
            return null;
        }

        private ImageSource BitmapToImageSource(System.Drawing.Bitmap bitmap)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                bitmap.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                ms.Position = 0;
                var bi = new BitmapImage();
                bi.BeginInit();
                bi.StreamSource = ms;
                bi.CacheOption = BitmapCacheOption.OnLoad;
                bi.EndInit();
                return bi;
            }
        }

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

        private void OnPathTypeChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!this.IsLoaded) return;
            SyncStyleFromUI();

            // ✨ 修复关键：使用 System.Windows.Visibility (类名) 而非属性名来访问 Visible/Collapsed
            if (SolidPanel != null)
            {
                SolidPanel.Visibility = (_currentStyle.PathType == PathCategory.Solid)
                    ? System.Windows.Visibility.Visible
                    : System.Windows.Visibility.Collapsed;
            }

            if (PatternPanel != null)
            {
                PatternPanel.Visibility = (_currentStyle.PathType == PathCategory.Pattern)
                    ? System.Windows.Visibility.Visible
                    : System.Windows.Visibility.Collapsed;
            }

            UpdatePreview();
        }



        /// <summary>
        /// 通用的参数变更事件处理。
        /// 合并了下拉框和滑块的事件，解决 CS0121 二义性错误。
        /// </summary>
        private void OnParamChanged(object sender, EventArgs e)
        {
            // 只有在窗口加载完成后才执行同步和预览，防止初始化时崩溃
            if (!this.IsLoaded) return;

            SyncStyleFromUI();
            UpdatePreview();
        }
        private void SelectColor_Click(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            var colorDlg = new Autodesk.AutoCAD.Windows.ColorDialog();
            colorDlg.Color = Autodesk.AutoCAD.Colors.Color.FromRgb(_currentStyle.MainColor.R, _currentStyle.MainColor.G, _currentStyle.MainColor.B);
            if (colorDlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                System.Drawing.Color c = colorDlg.Color.ColorValue;
                _currentStyle.MainColor = Color.FromRgb(c.R, c.G, c.B);
                if (ColorPreview != null) ColorPreview.Fill = new SolidColorBrush(_currentStyle.MainColor);
                UpdatePreview();
            }
        }

        private void StartDraw_Click(object sender, RoutedEventArgs e)
        {
            SyncStyleFromUI();
            this.Hide();
            try { GeometryEngine.DrawAnalysisLine(null, _currentStyle); }
            finally { this.Show(); }
        }

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
                        if (BlockLibraryCombo != null)
                        {
                            if (!BlockLibraryCombo.Items.Contains(br.Name)) BlockLibraryCombo.Items.Add(br.Name);
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
        private void GenerateLegend_Click(object sender, RoutedEventArgs e)
        {
            // 暂时隐藏窗口，方便 CAD 操作
            this.Hide();

            try
            {
                Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

                // 确保同步了当前的图层设置
                SyncStyleFromUI();

                using (doc.LockDocument())
                {
                    // 调用引擎生成图例
                    GeometryEngine.GenerateLegend(doc, _currentStyle.TargetLayer);
                }
            }
            catch (System.Exception ex)
            {
                Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog("生成图例出错: " + ex.Message);
            }
            finally
            {
                // 操作完成后显示回窗口
                this.Show();
            }
        }
        private void Close_Click(object sender, RoutedEventArgs e) => this.Close();
        /// <summary>
        /// 实现无边框窗口的鼠标拖动功能
        /// </summary>
        private void TitleBar_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            // 检查是否是鼠标左键按下
            if (e.LeftButton == System.Windows.Input.MouseButtonState.Pressed)
            {
                // 调用 WPF 窗口自带的拖动方法
                this.DragMove();
            }
        }
        /// <summary>
        /// 点击标题旁边的 ⓘ 按钮触发
        /// </summary>
        private void AboutBtn_Click(object sender, RoutedEventArgs e)
        {
            AboutWindow aboutWin = new AboutWindow();
            // 指定父窗口，确保 AboutWindow 浮动在主程序之上
            aboutWin.Owner = this;
            aboutWin.ShowDialog();
        }
    }
}