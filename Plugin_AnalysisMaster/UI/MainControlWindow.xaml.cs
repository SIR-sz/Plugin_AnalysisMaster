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

// 解决 CS0104：明确指定 UI 使用的类型别名
using Image = System.Windows.Controls.Image;
using Line = System.Windows.Shapes.Line;
using Path = System.IO.Path;

namespace Plugin_AnalysisMaster.UI
{
    public partial class MainControlWindow : Window
    {
        private AnalysisStyle _currentStyle = new AnalysisStyle();
        private static MainControlWindow _instance;

        // 用于释放非托管 HBitmap 资源
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

            // 1. 同步路径模式
            var typeItem = (ComboBoxItem)PathTypeCombo.SelectedItem;
            if (typeItem != null)
                _currentStyle.PathType = (PathCategory)Enum.Parse(typeof(PathCategory), typeItem.Tag.ToString());

            // 2. ✨ 修复：使用 SelectedItem 而非 Text，确保获取当前即时选中的块名
            // 因为在 SelectionChanged 事件触发瞬间，Text 属性可能还是旧值
            if (BlockLibraryCombo != null && BlockLibraryCombo.SelectedItem != null)
            {
                _currentStyle.SelectedBlockName = BlockLibraryCombo.SelectedItem.ToString();
            }
            else
            {
                _currentStyle.SelectedBlockName = BlockLibraryCombo?.Text ?? "";
            }

            // 3. 同步其他参数
            _currentStyle.Transparency = TransSlider?.Value ?? 0;
            _currentStyle.StartWidth = StartWidthSlider?.Value ?? 1.0;
            _currentStyle.MidWidth = MidWidthSlider?.Value ?? 0.8;
            _currentStyle.EndWidth = EndWidthSlider?.Value ?? 0.5;
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
                    System.Windows.Point pt1 = GetBezierPoint(i / 40.0, p1, p2, p3);
                    System.Windows.Point pt2 = GetBezierPoint((i + 1.0) / 40.0, p1, p2, p3);

                    // 此处使用 System.Windows.Shapes.Line (通过别名)
                    Line line = new Line
                    {
                        X1 = pt1.X,
                        Y1 = pt1.Y,
                        X2 = pt2.X,
                        Y2 = pt2.Y,
                        Stroke = brush,
                        StrokeThickness = 8,
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
                ImageSource imgSource = GetBlockImageSource(blockName);

                double step = 0.15;
                for (double t = 0; t <= 1; t += step)
                {
                    System.Windows.Point pt = GetBezierPoint(t, p1, p2, p3);
                    double imgSize = 22;

                    // 此处使用 System.Windows.Controls.Image (通过别名)
                    Image img = new Image
                    {
                        Source = imgSource,
                        Width = imgSize,
                        Height = imgSize,
                        Stretch = Stretch.Uniform
                    };

                    System.Windows.Point nPt = GetBezierPoint(Math.Min(t + 0.01, 1), p1, p2, p3);
                    Vector v = nPt - pt;
                    double angle = Math.Atan2(v.Y, v.X) * 180 / Math.PI;

                    img.RenderTransform = new RotateTransform(angle, imgSize / 2, imgSize / 2);
                    Canvas.SetLeft(img, pt.X - imgSize / 2);
                    Canvas.SetTop(img, pt.Y - imgSize / 2);
                    PreviewCanvas.Children.Add(img);
                }
            }
        }

        private void SyncBlockToCurrentDoc(string blockName)
        {
            var doc = AcApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            // ✨ 修复关键：从非模态窗口修改数据库必须先锁定文档，防止 eLockViolation
            using (DocumentLock loc = doc.LockDocument())
            {
                using (var tr = doc.Database.TransactionManager.StartTransaction())
                {
                    BlockTable bt = (BlockTable)tr.GetObject(doc.Database.BlockTableId, OpenMode.ForRead);
                    if (bt.Has(blockName)) return; // 如果已经导入过了，直接跳过

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
                                // 跨数据库克隆对象
                                doc.Database.WblockCloneObjects(ids, doc.Database.BlockTableId, mapping, DuplicateRecordCloning.Ignore, false);
                            }
                        }
                    }
                    tr.Commit();
                }
            }
        }

        private ImageSource GetBlockImageSource(string blockName)
        {
            var doc = AcApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return null;

            // 建议：由于是从 modeless 窗口访问，确保数据库处于可读状态
            using (var tr = doc.Database.TransactionManager.StartTransaction())
            {
                BlockTable bt = (BlockTable)tr.GetObject(doc.Database.BlockTableId, OpenMode.ForRead);
                if (bt.Has(blockName))
                {
                    var acColor = Autodesk.AutoCAD.Colors.Color.FromRgb(30, 30, 30);
                    IntPtr hBitmap = Autodesk.AutoCAD.Internal.Utils.GetBlockImage(bt[blockName], 128, 128, acColor);

                    if (hBitmap != IntPtr.Zero)
                    {
                        try
                        {
                            return Imaging.CreateBitmapSourceFromHBitmap(
                                hBitmap,
                                IntPtr.Zero,
                                Int32Rect.Empty,
                                BitmapSizeOptions.FromEmptyOptions());
                        }
                        finally
                        {
                            DeleteObject(hBitmap); // 必须释放非托管资源
                        }
                    }
                }
                tr.Commit();
            }
            return null;
        }

        private System.Windows.Point GetBezierPoint(double t, System.Windows.Point p0, System.Windows.Point p1, System.Windows.Point p2)
        {
            double x = (1 - t) * (1 - t) * p0.X + 2 * t * (1 - t) * p1.X + t * t * p2.X;
            double y = (1 - t) * (1 - t) * p0.Y + 2 * t * (1 - t) * p1.Y + t * t * p2.Y;
            return new System.Windows.Point(x, y);
        }

        private void OnPathTypeChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!this.IsLoaded) return;
            SyncStyleFromUI();

            // 解决 CS0176：显式使用 System.Windows.Visibility 类型名访问枚举成员
            if (SolidPanel != null)
                SolidPanel.Visibility = (_currentStyle.PathType == PathCategory.Solid)
                    ? System.Windows.Visibility.Visible : System.Windows.Visibility.Collapsed;

            if (PatternPanel != null)
                PatternPanel.Visibility = (_currentStyle.PathType == PathCategory.Pattern)
                    ? System.Windows.Visibility.Visible : System.Windows.Visibility.Collapsed;

            UpdatePreview();
        }

        private void OnParamChanged(object sender, SelectionChangedEventArgs e)
        {
            // 确保窗口加载完成后再执行，防止初始化时的空引用
            if (this.IsLoaded)
            {
                SyncStyleFromUI(); // 此时 SyncStyleFromUI 会通过 SelectedItem 拿到正确名字
                UpdatePreview();   // UpdatePreview 接着会根据正确名字去抓取 CAD 图片
            }
        }

        private void OnParamChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (this.IsLoaded) SyncStyleFromUI();
        }

        private void SelectColor_Click(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            Autodesk.AutoCAD.Windows.ColorDialog colorDlg = new Autodesk.AutoCAD.Windows.ColorDialog();
            colorDlg.Color = Autodesk.AutoCAD.Colors.Color.FromRgb(_currentStyle.MainColor.R, _currentStyle.MainColor.G, _currentStyle.MainColor.B);

            if (colorDlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                System.Drawing.Color selectedColor = colorDlg.Color.ColorValue;
                _currentStyle.MainColor = Color.FromRgb(selectedColor.R, selectedColor.G, selectedColor.B);
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