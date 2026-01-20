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
        /// 加载外部 DWG 资源库并根据命名约定填充下拉列表。
        /// 逻辑：Cap_ 开头的进入端头列表，Pat_ 开头的进入阵列列表。
        /// 优化：UI 只显示去掉前缀后的名称，真实名称隐藏在 Tag 中；默认添加“无”选项。
        /// </summary>
        private void LoadPatternLibrary()
        {
            if (BlockLibraryCombo == null || StartArrowCombo == null || EndArrowCombo == null) return;

            // 1. 初始化下拉框
            BlockLibraryCombo.Items.Clear();
            StartArrowCombo.Items.Clear();
            EndArrowCombo.Items.Clear();

            // 2. 预设“无”选项
            StartArrowCombo.Items.Add(new ComboBoxItem { Content = "无", Tag = "None" });
            EndArrowCombo.Items.Add(new ComboBoxItem { Content = "无", Tag = "None" });
            StartArrowCombo.SelectedIndex = 0; // 起始端默认选“无”
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
                            // 3. 根据前缀分发到不同的下拉框
                            if (realName.StartsWith("Cap_", StringComparison.OrdinalIgnoreCase))
                            {
                                string displayName = realName.Substring(4); // 去掉 Cap_
                                StartArrowCombo.Items.Add(new ComboBoxItem { Content = displayName, Tag = realName });
                                EndArrowCombo.Items.Add(new ComboBoxItem { Content = displayName, Tag = realName });
                            }
                            else if (realName.StartsWith("Pat_", StringComparison.OrdinalIgnoreCase))
                            {
                                string displayName = realName.Substring(4); // 去掉 Pat_
                                BlockLibraryCombo.Items.Add(new ComboBoxItem { Content = displayName, Tag = realName });
                            }
                        }
                    }
                }
            }
            if (BlockLibraryCombo.Items.Count > 0) BlockLibraryCombo.SelectedIndex = 0;
        }

        /// <summary>
        /// 将 UI 界面上的控件状态同步到静态模型 _currentStyle 中。
        /// 修改逻辑：增加了对 SkeletonTypeCombo 的读取，用于同步 IsCurved 属性（样条曲线 vs 多段线）。
        /// </summary>
        private void SyncStyleFromUI()
        {
            if (PathTypeCombo == null || _currentStyle == null) return;

            // 1. 同步渲染模式
            if (PathTypeCombo.SelectedItem is ComboBoxItem typeItem)
                _currentStyle.PathType = (PathCategory)Enum.Parse(typeof(PathCategory), typeItem.Tag.ToString());

            // 2. ✨ 同步骨架类型 (IsCurved)
            if (SkeletonTypeCombo != null && SkeletonTypeCombo.SelectedItem is ComboBoxItem skelItem)
            {
                _currentStyle.IsCurved = bool.Parse(skelItem.Tag.ToString());
            }

            // 3. 同步图块名称
            if (BlockLibraryCombo != null && BlockLibraryCombo.SelectedItem is ComboBoxItem patItem)
                _currentStyle.SelectedBlockName = patItem.Tag.ToString();

            if (StartArrowCombo != null && StartArrowCombo.SelectedItem is ComboBoxItem startItem)
                _currentStyle.StartArrowType = startItem.Tag.ToString();

            if (EndArrowCombo != null && EndArrowCombo.SelectedItem is ComboBoxItem endItem)
                _currentStyle.EndArrowType = endItem.Tag.ToString();

            // 4. 同步滑块参数
            _currentStyle.StartWidth = StartWidthSlider?.Value ?? 1.0;
            _currentStyle.MidWidth = MidWidthSlider?.Value ?? 0.8;
            _currentStyle.EndWidth = EndWidthSlider?.Value ?? 0.5;
            _currentStyle.PatternSpacing = SpacingSlider?.Value ?? 10.0;
            _currentStyle.PatternScale = PatternScaleSlider?.Value ?? 1.0;
            _currentStyle.ArrowSize = ArrowSizeSlider?.Value ?? 8.0;

            if (CapIndentSlider != null)
            {
                _currentStyle.CapIndent = CapIndentSlider.Value;
            }
        }

        /// <summary>
        /// 刷新预览画布。
        /// 修改逻辑：删除了 brush.Opacity 的计算逻辑，预览图强制以 1.0 (不透明) 显示，
        /// 并确保预览图依然受 visualScale 全局缩放保护。
        /// </summary>
        private void UpdatePreview()
        {
            if (PreviewCanvas == null || !this.IsLoaded) return;
            PreviewCanvas.Children.Clear();

            double w = PreviewCanvas.ActualWidth;
            double h = PreviewCanvas.ActualHeight;
            if (w <= 0 || h <= 0) return;

            System.Windows.Point p1 = new System.Windows.Point(w * 0.25, h * 0.7);
            System.Windows.Point p2 = new System.Windows.Point(w * 0.5, h * 0.2);
            System.Windows.Point p3 = new System.Windows.Point(w * 0.75, h * 0.5);

            // 计算全局视觉缩放系数 (防止溢出)
            double maxW = Math.Max(_currentStyle.StartWidth, Math.Max(_currentStyle.MidWidth, _currentStyle.EndWidth));
            double maxA = (_currentStyle.ArrowSize / 8.0) * 24.0;
            double maxP = _currentStyle.PatternScale * 22.0;
            double currentMax = Math.Max(maxW, Math.Max(maxA, maxP));
            double visualScale = currentMax > 40.0 ? 40.0 / currentMax : 1.0;

            // ✨ 修改：不再读取 Transparency，直接使用完全不透明的画刷
            var brush = new SolidColorBrush(_currentStyle.MainColor);

            double totalLenGuess = 200.0;
            double indentOffset = (_currentStyle.CapIndent * visualScale / totalLenGuess) * 0.5;
            indentOffset = Math.Max(-0.2, Math.Min(0.4, indentOffset));

            if (_currentStyle.PathType == PathCategory.Solid)
            {
                int segments = 40;
                double tStart = Math.Max(0, indentOffset);
                double tEnd = Math.Min(1, 1 - indentOffset);

                for (int i = 0; i < segments; i++)
                {
                    double t1 = tStart + (tEnd - tStart) * (i / (double)segments);
                    double t2 = tStart + (tEnd - tStart) * ((i + 1.0) / segments);
                    System.Windows.Point pt1 = GetBezierPoint(t1, p1, p2, p3);
                    System.Windows.Point pt2 = GetBezierPoint(t2, p1, p2, p3);
                    double thickness = CalculateBezierWidth(t1, _currentStyle.StartWidth, _currentStyle.MidWidth, _currentStyle.EndWidth) * visualScale;

                    Line line = new Line { X1 = pt1.X, Y1 = pt1.Y, X2 = pt2.X, Y2 = pt2.Y, Stroke = brush, StrokeThickness = thickness, StrokeStartLineCap = PenLineCap.Round, StrokeEndLineCap = PenLineCap.Round };
                    PreviewCanvas.Children.Add(line);
                }
            }
            else if (_currentStyle.PathType == PathCategory.Pattern && !string.IsNullOrEmpty(_currentStyle.SelectedBlockName))
            {
                SyncBlockToCurrentDoc(_currentStyle.SelectedBlockName);
                ImageSource mask = GetBlockMaskSource(_currentStyle.SelectedBlockName);
                double tStart = Math.Max(0, indentOffset);
                double tEnd = Math.Min(1, 1 - indentOffset);
                double step = 0.12;
                for (double t = tStart; t <= tEnd; t += step)
                {
                    System.Windows.Point pt = GetBezierPoint(t, p1, p2, p3);
                    double visualSize = (22 * _currentStyle.PatternScale) * visualScale;
                    RenderPreviewItem(pt, t, p1, p2, p3, mask, brush, visualSize);
                }
            }

            RenderArrowPreview(p1, 0.0, p1, p2, p3, _currentStyle.StartArrowType, brush, true, visualScale);
            RenderArrowPreview(p3, 1.0, p1, p2, p3, _currentStyle.EndArrowType, brush, false, visualScale);
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
        /// 渲染预览图中的单个图块项。
        /// 此方法已在 UpdatePreview 中通过传入带有 visualScale 修正后的 size 实现了大小控制。
        /// </summary>
        private void RenderPreviewItem(System.Windows.Point pt, double t, System.Windows.Point p0, System.Windows.Point p1, System.Windows.Point p2, ImageSource mask, Brush brush, double size)
        {
            if (mask == null) return;

            // 安全下限：防止阵列单元太小不可见
            if (size < 2) size = 2;

            var rect = new System.Windows.Shapes.Rectangle { Width = size, Height = size, Fill = brush, OpacityMask = new ImageBrush(mask) };

            System.Windows.Point nPt = GetBezierPoint(Math.Min(t + 0.01, 1), p0, p1, p2);
            Vector v = nPt - pt;
            double angle = Math.Atan2(v.Y, v.X) * 180 / Math.PI;

            rect.RenderTransform = new RotateTransform(angle, size / 2, size / 2);
            Canvas.SetLeft(rect, pt.X - size / 2);
            Canvas.SetTop(rect, pt.Y - size / 2);
            PreviewCanvas.Children.Add(rect);
        }

        /// <summary>
        /// 渲染起终点箭头的预览。
        /// 修改逻辑：切线计算现在会根据 IsCurved 模式自动切换。
        /// 在折线模式下，箭头的朝向会严格锁定在第一段或最后一段线段的方向上。
        /// </summary>
        private void RenderArrowPreview(System.Windows.Point pt, double t, System.Windows.Point p0, System.Windows.Point p1, System.Windows.Point p2, string blockName, Brush brush, bool isStart, double visualScale)
        {
            if (string.IsNullOrEmpty(blockName) || blockName == "None") return;

            SyncBlockToCurrentDoc(blockName);
            ImageSource mask = GetBlockMaskSource(blockName);
            if (mask == null) return;

            double size = 24 * (_currentStyle.ArrowSize / 8.0) * visualScale;
            if (size < 4) size = 4;

            var rect = new System.Windows.Shapes.Rectangle { Width = size, Height = size, Fill = brush, OpacityMask = new ImageBrush(mask) };

            // ✨ 计算切线：根据是否为曲线模式采用不同算法
            Vector v;
            if (_currentStyle.IsCurved)
            {
                System.Windows.Point nextPt = GetBezierPoint(isStart ? 0.01 : 1.0, p0, p1, p2);
                System.Windows.Point prevPt = GetBezierPoint(isStart ? 0.0 : 0.99, p0, p1, p2);
                v = nextPt - prevPt;
            }
            else
            {
                // 折线模式直接取第一段 (p0-p1) 或最后一段 (p1-p2) 的方向
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

        private void OnParamChanged(object sender, SelectionChangedEventArgs e)
        {
            if (this.IsLoaded) { SyncStyleFromUI(); UpdatePreview(); }
        }

        /// <summary>
        /// 当滑块类控件（Slider）的数值发生变化时触发的事件回调。
        /// 该方法负责在用户拖动滑块（如调整宽度、透明度、间距等）时，
        /// 实时将 UI 的最新数值同步到内存模型，并立即触发预览图的重绘，确保交互反馈的实时性。
        /// </summary>
        private void OnParamChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (this.IsLoaded)
            {
                // 1. 同步最新的滑块数值到 _currentStyle 模型
                SyncStyleFromUI();

                // 2. 核心修复：立即触发预览 Canvas 的重绘逻辑
                UpdatePreview();
            }
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

        private void Close_Click(object sender, RoutedEventArgs e) => this.Close();
    }
}