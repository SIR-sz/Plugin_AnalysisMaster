using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Plugin_AnalysisMaster.Models;
using Plugin_AnalysisMaster.Services;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media;

namespace Plugin_AnalysisMaster.UI
{
    public partial class AnimationWindow : Window
    {
        // 使用 ObservableCollection 确保 UI 列表自动刷新
        private ObservableCollection<AnimPathItem> _pathList = new ObservableCollection<AnimPathItem>();
        private CancellationTokenSource _cts;
        // 在 AnimationWindow 类中添加以下方法
        /// <summary>
        /// 当窗口失去焦点（例如用户点击了 AutoCAD 绘图区或其他窗口）时触发。
        /// 作用：自动取消当前列表的选中状态并擦除屏幕上的高亮骨架线，确保“去忙别的”时图形自动清理。
        /// </summary>
        private void Window_Deactivated(object sender, EventArgs e)
        {
            // 强制取消选中并清理高亮
            if (PathListView != null)
            {
                PathListView.SelectedItem = null;
            }
            GeometryEngine.ClearHighlightTransient();
        }
        /// <summary>
        /// 在 ListView 内部按下鼠标时的预览处理。
        /// 作用：通过 VisualTreeHelper 判定点击位置。如果点击的是列表内的空白处（没有点击到行），则执行取消选中。
        /// </summary>
        private void PathListView_PreviewMouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            // 检查点击的目标是否属于 ListViewItem（即具体的行）
            DependencyObject dep = (DependencyObject)e.OriginalSource;
            while (dep != null && !(dep is ListViewItem))
            {
                dep = VisualTreeHelper.GetParent(dep);
            }

            // 如果 dep 为 null，说明点击的是列表背景空白处
            if (dep == null)
            {
                PathListView.SelectedItem = null;
                GeometryEngine.ClearHighlightTransient();
            }
        }
        /// <summary>
        /// 窗口根级容器的鼠标按下预览。
        /// 作用：实现“点击窗口内非列表区域（如按钮边距、标题栏空白处等）”即取消高亮的逻辑。
        /// </summary>
        private void RootBorder_PreviewMouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            // 检查点击点是否在 ListView 控件的几何范围之内
            var hitTestResult = VisualTreeHelper.HitTest(PathListView, e.GetPosition(PathListView));

            // 如果点击点完全不在 ListView 内部（例如点击了窗口侧边空白或按钮间隔），则取消选中
            if (hitTestResult == null)
            {
                PathListView.SelectedItem = null;
                GeometryEngine.ClearHighlightTransient();
            }
        }
        private void TitleBar_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (e.LeftButton == System.Windows.Input.MouseButtonState.Pressed)
            {
                this.DragMove(); // 允许点击标题栏拖动窗口
            }
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            this.Close(); // 关闭窗口
        }
        public AnimationWindow()
        {
            InitializeComponent();
            PathListView.ItemsSource = _pathList;
        }

        #region 1. 列表操作 (添加、移除、清空)

        /// <summary>
        /// 添加路径按钮点击事件。
        /// 修改说明：将去重逻辑由“指纹识别”改为“实体 ID (ObjectId) 识别”，从而支持将样式完全相同的不同线条同时添加到列表中。
        /// </summary>
        private void AddPath_Click(object sender, RoutedEventArgs e)
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            Editor ed = doc.Editor;

            this.Visibility = System.Windows.Visibility.Collapsed;
            try
            {
                var res = ed.GetSelection();
                if (res.Status != PromptStatus.OK) return;

                using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
                {
                    // ✨ 修改：改用 ObjectId 字符串作为本次选择的去重 key
                    HashSet<string> processedInThisBatch = new HashSet<string>();
                    int addedCount = 0;

                    foreach (SelectedObject selObj in res.Value)
                    {
                        ObjectId actualDataHolder;
                        string fingerprint = GetAnimFingerprint(tr, selObj.ObjectId, out actualDataHolder);

                        if (string.IsNullOrEmpty(fingerprint)) continue;

                        // ✨ 核心修改：检查实体 ID 是否已在列表中，而不是检查样式是否重复
                        if (processedInThisBatch.Contains(actualDataHolder.ToString()) || IsPathAlreadyExists(actualDataHolder))
                            continue;

                        processedInThisBatch.Add(actualDataHolder.ToString());

                        Entity ent = tr.GetObject(actualDataHolder, OpenMode.ForRead) as Entity;
                        var item = CreatePathItemFromEntity(tr, ent, fingerprint);
                        if (item != null)
                        {
                            _pathList.Add(item);
                            addedCount++;
                        }
                    }
                    tr.Commit();
                    ed.WriteMessage($"\n[成功] 已识别并添加 {addedCount} 条分析路径。");
                }
            }
            finally
            {
                this.Visibility = System.Windows.Visibility.Visible;
            }
        }

        /// <summary>
        /// 辅助方法：检查列表中是否已存在指定的实体。
        /// 修改说明：原名为 IsFingerprintAlreadyExists，现根据 ID 检查唯一性，确保不重复添加同一个实体。
        /// </summary>
        private bool IsPathAlreadyExists(ObjectId id)
        {
            // 直接在 ObservableCollection 中查找是否存在相同的 ObjectId
            return _pathList.Any(x => x.Id == id);
        }

        // 辅助方法：获取实体的动画指纹
        // 修改位置：AnimationWindow.xaml.cs 约 132 行
        // ✨ 新增 out 参数，用于返回真正持有数据的 ID
        private string GetAnimFingerprint(Transaction tr, ObjectId id, out ObjectId dataHolderId)
        {
            dataHolderId = ObjectId.Null;
            if (id.IsNull || !id.IsValid || id.IsErased) return "";

            // 1. 检查当前物体
            string fp = GetDirectFingerprint(tr, id);
            if (!string.IsNullOrEmpty(fp))
            {
                dataHolderId = id;
                return fp;
            }

            // 2. 检查所属编组
            Entity ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
            if (ent != null)
            {
                ObjectIdCollection reactors = ent.GetPersistentReactorIds();
                if (reactors != null)
                {
                    foreach (ObjectId rId in reactors)
                    {
                        if (rId.IsValid && rId.ObjectClass.IsDerivedFrom(RXClass.GetClass(typeof(Group))))
                        {
                            Group gp = tr.GetObject(rId, OpenMode.ForRead) as Group;
                            if (gp != null)
                            {
                                foreach (ObjectId memberId in gp.GetAllEntityIds())
                                {
                                    string memberFp = GetDirectFingerprint(tr, memberId);
                                    if (!string.IsNullOrEmpty(memberFp))
                                    {
                                        dataHolderId = memberId; // ✨ 记录持有数据的成员
                                        return memberFp;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            return "";
        }
        private string GetDirectFingerprint(Transaction tr, ObjectId id)
        {
            Entity ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
            if (ent == null || ent.ExtensionDictionary.IsNull) return "";

            DBDictionary dict = tr.GetObject(ent.ExtensionDictionary, OpenMode.ForRead) as DBDictionary;
            if (!dict.Contains("ANALYSIS_ANIM_DATA")) return "";

            Xrecord xRec = tr.GetObject(dict.GetAt("ANALYSIS_ANIM_DATA"), OpenMode.ForRead) as Xrecord;
            using (ResultBuffer rb = xRec.Data)
            {
                if (rb == null) return "";
                var arr = rb.AsArray();
                return arr.Length > 0 ? arr[0].Value.ToString() : "";
            }
        }
        private void RemovePath_Click(object sender, RoutedEventArgs e)
        {
            var selectedItems = PathListView.SelectedItems.Cast<AnimPathItem>().ToList();
            foreach (var item in selectedItems)
            {
                _pathList.Remove(item);
            }
        }

        private void ClearPaths_Click(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show("确定要清空播放列表吗？", "提示", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
            {
                _pathList.Clear();
            }
        }

        // 文件：AnimationWindow.xaml.cs

        /// <summary>
        /// 根据实体和指纹数据创建列表项对象。
        /// 修改说明：
        /// 1. 移除了图层显示：不再使用 "图层: LayerName" 格式。
        /// 2. 描述同步图例：解析指纹数据，根据 PathType 生成“连续线样式”或“阵列样式(图块名)”描述。
        /// </summary>
        private AnimPathItem CreatePathItemFromEntity(Transaction tr, Entity ent, string fingerprint)
        {
            string[] parts = fingerprint.Split('|');
            if (parts.Length < 6) return null;

            try
            {
                // ✨ 解析样式描述（同步 GeometryEngine.GenerateLegend 的逻辑）
                var pathType = parts[0]; // PathCategory 字符串
                var blockName = parts[2]; // SelectedBlockName

                string description = (pathType == "Solid")
                    ? "连续线样式"
                    : $"阵列样式({(string.IsNullOrEmpty(blockName) ? "默认" : blockName)})";

                var colorStr = parts[5];
                var mediaColor = (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(colorStr);

                return new AnimPathItem
                {
                    Id = ent.ObjectId,
                    Name = description, // ✨ 使用新的描述逻辑
                    PathColor = mediaColor,
                    GroupNumber = 1
                };
            }
            catch
            {
                return new AnimPathItem { Id = ent.ObjectId, Name = "未知样式", PathColor = Colors.SteelBlue, GroupNumber = 1 };
            }
        }

        #endregion

        #region 2. 播放控制

        /// <summary>
        /// 播放按钮点击事件处理逻辑。
        /// 修改说明：将 speed 的变量类型从 int 改为 double，以支持 0.1 - 10 倍速的浮点数调节，
        /// 并同步修改了传给 GeometryEngine.PlaySequenceAsync 的参数类型。
        /// </summary>
        private async void BtnPlay_Click(object sender, RoutedEventArgs e)
        {
            if (_pathList.Count == 0) return;

            BtnPlay.IsEnabled = false;
            BtnStop.IsEnabled = true;
            _cts = new CancellationTokenSource();

            try
            {
                double thickness = ThicknessSlider.Value;
                // ✨ 修改：由 (int) 改为直接获取 double 值
                double speed = SpeedSlider.Value;
                bool isLoop = LoopCheck.IsChecked == true;
                bool isPersistent = PersistenceCheck.IsChecked == true;

                await GeometryEngine.PlaySequenceAsync(
                    _pathList.ToList(),
                    thickness,
                    speed, // 传入 double 类型倍速
                    isLoop,
                    isPersistent,
                    _cts.Token);
            }
            catch (OperationCanceledException) { }
            catch (System.Exception ex) { MessageBox.Show("播放出错: " + ex.Message); }
            finally
            {
                BtnPlay.IsEnabled = true;
                BtnStop.IsEnabled = false;
                _cts?.Dispose();
                _cts = null;
            }
        }

        private void BtnStop_Click(object sender, RoutedEventArgs e)
        {
            _cts?.Cancel();
        }

        #endregion

        #region 3. UI 交互逻辑 (编号提示等)
        // 1. 响应“显示编号”按钮点击
        private void BtnShowLabels_Click(object sender, RoutedEventArgs e)
        {
            bool isShow = BtnShowLabels.IsChecked == true;
            GeometryEngine.UpdatePathLabels(_pathList.ToList(), isShow);
        }

        // 2. 列表选中高亮逻辑
        private ObjectId _lastSelectedId = ObjectId.Null;
        /// <summary>
        /// 列表选中项变更事件处理。
        /// 修改说明：简化了逻辑，移除了对 _lastSelectedId 的依赖。
        /// 现在由 GeometryEngine 内部处理“先清理、后绘制”的逻辑，确保了切换的流畅性和准确性。
        /// </summary>
        private void PathListView_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // 如果选中了有效条目
            if (PathListView.SelectedItem is AnimPathItem item)
            {
                // 直接触发高亮（内部会自动先清理旧线）
                GeometryEngine.HighlightPath(item.Id, true);
            }
            else
            {
                // ✨ 核心修复：如果用户点击空白处取消了选中，确保立即擦除屏幕上的高亮线
                GeometryEngine.ClearHighlightTransient();
            }
        }
        // 3. 确保窗口关闭时清理标签
        /// <summary>
        /// 确保窗口关闭时清理所有瞬态图形（包括路径编号和高亮线条）。
        /// 修改说明：增加了对 ClearHighlightTransient 的调用，防止关闭动画管理窗口后，高亮骨架线残留在图中。
        /// </summary>
        protected override void OnClosed(EventArgs e)
        {
            // 清理路径编号标签
            GeometryEngine.ClearPathLabels();

            // ✨ 新增：清理高亮瞬态骨架线
            GeometryEngine.ClearHighlightTransient();

            base.OnClosed(e);
        }
        private void ShowNumberTip(ObjectId id, int index)
        {
            // 这里可以调用 GeometryEngine 里的一个简单方法
            // 在 CAD 中对该 ID 对应的物体进行闪烁或显示瞬态编号
            // 篇幅原因，建议在 GeometryEngine 中实现，此处仅作为调用锚点
        }

        #endregion
    }

    #region 4. 值转换器 (用于 XAML 绑定)

    // 如果你没有在独立文件中定义，可以暂时放在这里
    public class ColorToBrushConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if (value is System.Windows.Media.Color color)
                return new SolidColorBrush(color);
            return Brushes.Gray;
        }
        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture) => throw new NotImplementedException();
    }

    public class IndexConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            ListViewItem item = (ListViewItem)value;
            ListView listView = ItemsControl.ItemsControlFromItemContainer(item) as ListView;
            int index = listView.ItemContainerGenerator.IndexFromContainer(item);
            return (index + 1).ToString();
        }
        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture) => throw new NotImplementedException();
    }

    #endregion
}