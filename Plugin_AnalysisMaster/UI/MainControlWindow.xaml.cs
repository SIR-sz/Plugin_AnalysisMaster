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
        private AnalysisStyle _currentStyle = new AnalysisStyle();
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

        private void SyncStyleFromUI()
        {
            if (PathTypeCombo == null || _currentStyle == null) return;

            // 1. 同步路径模式 (实线 vs 阵列)
            var typeItem = (ComboBoxItem)PathTypeCombo.SelectedItem;
            if (typeItem != null)
                _currentStyle.PathType = (PathCategory)Enum.Parse(typeof(PathCategory), typeItem.Tag.ToString());

            // 2. 同步选中的图块名称 (用于阵列模式)
            if (BlockLibraryCombo != null && BlockLibraryCombo.SelectedItem != null)
                _currentStyle.SelectedBlockName = BlockLibraryCombo.SelectedItem.ToString();

            // 3. ✨ 新增：同步起点箭头类型
            if (StartArrowCombo != null && StartArrowCombo.SelectedItem != null)
            {
                // 获取选中的文本（如果是 ComboBoxItem 则取 Content）
                string val = (StartArrowCombo.SelectedItem is ComboBoxItem cbi)
                             ? cbi.Content.ToString()
                             : StartArrowCombo.SelectedItem.ToString();

                // 如果选的是“无”，则存入 "None"，否则存入块名
                _currentStyle.StartArrowType = (val == "无" || val == "None") ? "None" : val;
            }

            // 4. ✨ 新增：同步结束端箭头类型
            if (EndArrowCombo != null && EndArrowCombo.SelectedItem != null)
            {
                string val = (EndArrowCombo.SelectedItem is ComboBoxItem cbi)
                             ? cbi.Content.ToString()
                             : EndArrowCombo.SelectedItem.ToString();

                _currentStyle.EndArrowType = (val == "无" || val == "None") ? "None" : val;
            }

            // 5. 同步颜色透明度
            _currentStyle.Transparency = TransSlider?.Value ?? 0;

            // 6. 同步实线宽度参数 (起点、中点、终点)
            _currentStyle.StartWidth = StartWidthSlider?.Value ?? 1.0;
            _currentStyle.MidWidth = MidWidthSlider?.Value ?? 0.8;
            _currentStyle.EndWidth = EndWidthSlider?.Value ?? 0.5;

            // 7. 同步排列间距
            if (SpacingSlider != null)
            {
                _currentStyle.PatternSpacing = SpacingSlider.Value;
            }

            // 8. 同步图块缩放比例
            _currentStyle.PatternScale = PatternScaleSlider?.Value ?? 1.0;
        }

        private void UpdatePreview()
        {
            if (PreviewCanvas == null || !this.IsLoaded) return;
            PreviewCanvas.Children.Clear();

            double w = PreviewCanvas.ActualWidth;
            double h = PreviewCanvas.ActualHeight;
            if (w <= 0 || h <= 0) return;

            System.Windows.Point p1 = new System.Windows.Point(w * 0.1, h * 0.75);
            System.Windows.Point p2 = new System.Windows.Point(w * 0.5, h * 0.1);
            System.Windows.Point p3 = new System.Windows.Point(w * 0.9, h * 0.5);

            var brush = new SolidColorBrush(_currentStyle.MainColor);
            brush.Opacity = (100 - _currentStyle.Transparency) / 100.0;

            if (_currentStyle.PathType == PathCategory.Solid)
            {
                int segments = 40;
                for (int i = 0; i < segments; i++)
                {
                    double t = i / 40.0;
                    System.Windows.Point pt1 = GetBezierPoint(t, p1, p2, p3);
                    System.Windows.Point pt2 = GetBezierPoint((i + 1.0) / 40.0, p1, p2, p3);

                    // 计算真实宽度
                    double thickness = CalculateBezierWidth(t, _currentStyle.StartWidth, _currentStyle.MidWidth, _currentStyle.EndWidth);

                    // ✨ 修复：视觉压缩逻辑。即使 CAD 里的宽度是 10000，预览里的最大厚度也限制在 40 像素，防止溢出
                    double previewThickness = Math.Min(thickness * 0.05, 40);
                    if (previewThickness < 1) previewThickness = 1;

                    Line line = new Line
                    {
                        X1 = pt1.X,
                        Y1 = pt1.Y,
                        X2 = pt2.X,
                        Y2 = pt2.Y,
                        Stroke = brush,
                        StrokeThickness = previewThickness,
                        StrokeStartLineCap = PenLineCap.Round,
                        StrokeEndLineCap = PenLineCap.Round
                    };
                    PreviewCanvas.Children.Add(line);
                }
            }
            else if (_currentStyle.PathType == PathCategory.Pattern)
            {
                string blockName = _currentStyle.SelectedBlockName;
                if (string.IsNullOrEmpty(blockName)) return;

                SyncBlockToCurrentDoc(blockName);
                ImageSource maskSource = GetBlockMaskSource(blockName);

                double step = 0.15;
                for (double t = 0; t <= 1; t += step)
                {
                    System.Windows.Point pt = GetBezierPoint(t, p1, p2, p3);
                    double imgSize = 22;

                    var colorRect = new System.Windows.Shapes.Rectangle
                    {
                        Width = imgSize,
                        Height = imgSize,
                        Fill = brush
                    };

                    if (maskSource != null)
                        colorRect.OpacityMask = new ImageBrush(maskSource);

                    System.Windows.Point nPt = GetBezierPoint(Math.Min(t + 0.01, 1), p1, p2, p3);
                    Vector v = nPt - pt;
                    double angle = Math.Atan2(v.Y, v.X) * 180 / Math.PI;

                    colorRect.RenderTransform = new RotateTransform(angle, imgSize / 2, imgSize / 2);
                    Canvas.SetLeft(colorRect, pt.X - imgSize / 2);
                    Canvas.SetTop(colorRect, pt.Y - imgSize / 2);
                    PreviewCanvas.Children.Add(colorRect);
                }
            }
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

        private void OnParamChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (this.IsLoaded) SyncStyleFromUI();
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