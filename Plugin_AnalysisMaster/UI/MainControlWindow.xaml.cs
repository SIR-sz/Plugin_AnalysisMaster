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
            if (PathTypeCombo == null) return;
            var item = (ComboBoxItem)PathTypeCombo.SelectedItem;
            _currentStyle.PathType = (PathCategory)Enum.Parse(typeof(PathCategory), item.Tag.ToString());

            // 物理宽度
            _currentStyle.StartWidth = StartWidthSlider?.Value ?? 1.0;
            _currentStyle.MidWidth = MidWidthSlider?.Value ?? 0.8;
            _currentStyle.EndWidth = EndWidthSlider?.Value ?? 0.5;

            // 阵列参数
            _currentStyle.SelectedBlockName = BlockLibraryCombo?.SelectedItem?.ToString() ?? "";
            _currentStyle.PatternSpacing = SpacingSlider?.Value ?? 10.0;
            _currentStyle.PatternScale = PatternScaleSlider?.Value ?? 1.0;

            // 端头
            _currentStyle.ArrowSize = ArrowSizeSlider?.Value ?? 8.0;
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

        private void UpdatePreview() { /* 保持 Canvas 预览逻辑 */ }
        private void OnParamChanged(object sender, RoutedPropertyChangedEventArgs<double> e) { if (this.IsLoaded) UpdatePreview(); }
        private void OnParamChanged(object sender, SelectionChangedEventArgs e) { if (this.IsLoaded) UpdatePreview(); }
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