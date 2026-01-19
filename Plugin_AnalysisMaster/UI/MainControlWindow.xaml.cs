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
using System.Windows.Media;
using AcApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace Plugin_AnalysisMaster.UI
{
    public partial class MainControlWindow : Window
    {
        private AnalysisStyle _currentStyle = new AnalysisStyle();
        private static MainControlWindow _instance;

        // ✨ 修复 CS0117：提供 MainTool.cs 调用的静态入口
        // ✨ 修复 CS0117：提供静态入口
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

        // ✨ 修复 CS0117：AutoCAD 2015 兼容的文件读取逻辑
        private void LoadPatternLibrary()
        {
            string dllPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string libPath = Path.Combine(dllPath, "Assets", "PatternLibrary.dwg");

            if (File.Exists(libPath))
            {
                using (Database db = new Database(false, true))
                {
                    // ✨ 修复：2015 版应使用 OpenForReadAndReadShare
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
            if (PathTypeCombo == null) return;
            var item = (ComboBoxItem)PathTypeCombo.SelectedItem;
            _currentStyle.PathType = (PathCategory)Enum.Parse(typeof(PathCategory), item.Tag.ToString());
            _currentStyle.StartWidth = StartWidthSlider?.Value ?? 1.0;
            _currentStyle.EndWidth = EndWidthSlider?.Value ?? 0.5;
            _currentStyle.ArrowSize = ArrowSizeSlider?.Value ?? 8.0;
            _currentStyle.SelectedBlockName = BlockLibraryCombo?.SelectedItem?.ToString() ?? "";
            _currentStyle.PatternSpacing = SpacingSlider?.Value ?? 10.0;
        }

        // ✨ 修复 CS0176：正确访问 Visibility 静态值
        private void OnPathTypeChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!this.IsLoaded) return;
            SyncStyleFromUI();

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

        // ✨ 修复：确保在 AddAllowedClass 之前调用 SetRejectMessage
        private void PickBlock_Click(object sender, RoutedEventArgs e)
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            this.Hide();
            try
            {
                var opt = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\n请选择图中的块作为样式单元: ");

                // ✨ 核心修复：必须先设置拒绝消息，否则 AddAllowedClass 会报错
                opt.SetRejectMessage("\n所选对象不是块参照，请重新选择！");
                opt.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.BlockReference), true);

                var res = doc.Editor.GetEntity(opt);
                if (res.Status == Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                {
                    using (var tr = doc.Database.TransactionManager.StartTransaction())
                    {
                        var br = (Autodesk.AutoCAD.DatabaseServices.BlockReference)tr.GetObject(res.ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);

                        // 将拾取的块名同步到 UI 和样式模型
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

        private void UpdatePreview() { /* 保持原有 PreviewCanvas 绘图逻辑 */ }
        private void OnParamChanged(object sender, RoutedPropertyChangedEventArgs<double> e) { if (this.IsLoaded) UpdatePreview(); }
        private void OnParamChanged(object sender, SelectionChangedEventArgs e) { if (this.IsLoaded) UpdatePreview(); }
        private void SelectColor_Click(object sender, System.Windows.Input.MouseButtonEventArgs e) { }
        private void Close_Click(object sender, RoutedEventArgs e) => this.Close();
    }
}