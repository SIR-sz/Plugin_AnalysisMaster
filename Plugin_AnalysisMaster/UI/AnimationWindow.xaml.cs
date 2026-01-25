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
        // 文件：Plugin_AnalysisMaster/UI/AnimationWindow.xaml.cs

        private void RootBorder_PreviewMouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            // ✨ 核心修复：如果点击的是按钮、输入框等交互控件，直接退出，不清理选中状态
            DependencyObject dep = (DependencyObject)e.OriginalSource;
            while (dep != null)
            {
                if (dep is System.Windows.Controls.Button || dep is System.Windows.Controls.Primitives.ToggleButton ||
                    dep is System.Windows.Controls.TextBox || dep is System.Windows.Controls.Slider)
                {
                    return;
                }
                dep = VisualTreeHelper.GetParent(dep);
            }

            // 检查点击点是否在 ListView 控件的几何范围之内
            var hitTestResult = VisualTreeHelper.HitTest(PathListView, e.GetPosition(PathListView));

            // 如果点击点完全不在 ListView 内部，则取消选中并清理高亮
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

            // ✨ 窗口加载时自动恢复
            this.Loaded += (s, e) => RestoreSequence();
        }
        /// <summary>
        /// 拦截关闭事件，改用隐藏方式，实现 Session 级持久化。
        /// </summary>
        protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
        {
            e.Cancel = true;
            this.Hide();
        }

        /// <summary>
        /// 从图纸数据库中恢复动画播放序列。
        /// 改进点：在创建 fullItem 后，显式将数据库读取的 Name 和样式属性赋值给 UI 模型。
        /// </summary>
        private void RestoreSequence()
        {
            var baseItems = GeometryEngine.LoadSequenceFromDwg();
            if (baseItems.Count == 0) return;

            _pathList.Clear();
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            bool hasGhostItems = false;
            using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
            {
                foreach (var baseItem in baseItems)
                {
                    if (baseItem.Id.IsNull || !baseItem.Id.IsValid || baseItem.Id.IsErased)
                    {
                        hasGhostItems = true;
                        continue;
                    }

                    string fp = GetAnimFingerprint(tr, baseItem.Id, out _);
                    if (!string.IsNullOrEmpty(fp))
                    {
                        Entity ent = tr.GetObject(baseItem.Id, OpenMode.ForRead) as Entity;
                        var fullItem = CreatePathItemFromEntity(tr, ent, fp);
                        if (fullItem != null)
                        {
                            // ✨ 核心修复：用从图纸读取的数据覆盖 UI 的默认生成值
                            fullItem.GroupNumber = baseItem.GroupNumber;
                            fullItem.LineStyle = baseItem.LineStyle;
                            if (!string.IsNullOrEmpty(baseItem.Name))
                            {
                                fullItem.Name = baseItem.Name; // 恢复自定义描述
                            }

                            _pathList.Add(fullItem);
                        }
                    }
                    else { hasGhostItems = true; }
                }
                tr.Commit();
            }

            if (hasGhostItems) GeometryEngine.SaveSequenceToDwg(_pathList);
        }
        /// <summary>
        /// 响应动画图例生成按钮点击。
        /// </summary>
        private void BtnLegend_Click(object sender, RoutedEventArgs e)
        {
            if (_pathList != null && _pathList.Count > 0)
            {
                GeometryEngine.GenerateAnimLegend(_pathList.ToList());
            }
        }

        /// <summary>
        /// 当路径描述编辑框失去焦点时，立即保存修改到图纸。
        /// </summary>
        private void Description_LostFocus(object sender, RoutedEventArgs e)
        {
            if (this.IsLoaded && _pathList != null)
            {
                GeometryEngine.SaveSequenceToDwg(_pathList);
            }
        }
        /// <summary>
        /// 响应线型下拉框变更事件。
        /// 作用：当用户在界面修改某条路径的实虚线设置时，立即触发持久化保存，防止数据丢失。
        /// </summary>
        private void LineStyle_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // 只有在窗口完全加载后（避免初始化触发）且列表有数据时才执行保存
            if (this.IsLoaded && _pathList != null)
            {
                GeometryEngine.SaveSequenceToDwg(_pathList);
            }
        }
        #region 1. 列表操作 (添加、移除、清空)

        // 文件位置：Plugin_AnalysisMaster/UI/AnimationWindow.xaml.cs
        // 约 103 行左右，#region 1. 列表操作 (添加、移除、清空) 内部

        /// <summary>
        /// 添加路径按钮点击事件。
        /// 修改说明：
        /// 1. 增加了有效性校验：通过 GetAnimFingerprint 过滤掉没有动画点位的线。
        /// 2. 增加了持久化调用：添加成功后自动执行 SaveSequenceToDwg。
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
                    HashSet<string> processedInThisBatch = new HashSet<string>();
                    int addedCount = 0;

                    foreach (SelectedObject selObj in res.Value)
                    {
                        ObjectId actualDataHolder;
                        // ✨ 这里的 GetAnimFingerprint 内部已经包含了“点数 > 2”的过滤逻辑
                        string fingerprint = GetAnimFingerprint(tr, selObj.ObjectId, out actualDataHolder);

                        // 如果指纹为空（说明没数据或点位不足），直接跳过
                        if (string.IsNullOrEmpty(fingerprint)) continue;

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

                    // ✨ 核心修改：如果添加了新项，立即持久化到图纸 NOD
                    if (addedCount > 0)
                    {
                        GeometryEngine.SaveSequenceToDwg(_pathList);
                        ed.WriteMessage($"\n[成功] 已识别并添加 {addedCount} 条有效分析路径。");
                    }
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

                // ✨ 过滤关键点：[1] 存储的是点数，如果点数 < 2，则视为无效动画数据
                if (arr.Length > 1 && arr[1].Value is int count && count < 2)
                    return "";

                return arr[0].Value.ToString();
            }
        }
        // 文件位置：Plugin_AnalysisMaster/UI/AnimationWindow.xaml.cs
        // 约 214 行左右

        /// <summary>
        /// 移除选中路径按钮。
        /// 修改说明：移除项后立即同步更新图纸中保存的序列数据。
        /// </summary>
        private void RemovePath_Click(object sender, RoutedEventArgs e)
        {
            var selectedItems = PathListView.SelectedItems.Cast<AnimPathItem>().ToList();
            if (selectedItems.Count == 0) return;

            foreach (var item in selectedItems)
            {
                _pathList.Remove(item);
            }

            // ✨ 核心修复：移除后必须立即保存到图纸，否则下次打开又会从 NOD 恢复
            GeometryEngine.SaveSequenceToDwg(_pathList);
        }

        private void ClearPaths_Click(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show("确定要清空播放列表吗？", "提示", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
            {
                _pathList.Clear();
                // ✨ 核心修复：清空后保存一个空列表到图纸，彻底清除持久化数据
                GeometryEngine.SaveSequenceToDwg(_pathList);
            }
        }
        private void GroupNumber_LostFocus(object sender, RoutedEventArgs e)
        {
            // 用户修改完组号并点击别处后，自动保存
            GeometryEngine.SaveSequenceToDwg(_pathList);
        }


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
        /// 响应“播放”按钮点击。
        /// 包含逻辑：参数获取、任务取消令牌管理、调用智能双模式引擎。
        /// </summary>
        private async void BtnPlay_Click(object sender, RoutedEventArgs e)
        {
            // 1. 如果当前正在播放，则先停止（防止任务堆叠）
            if (_cts != null)
            {
                _cts.Cancel();
                _cts.Dispose();
                _cts = null;
            }

            // 2. 检查是否有可播放的路径
            if (_pathList == null || _pathList.Count == 0)
            {
                return;
            }

            // 3. 从 UI 获取播放参数
            // 注意：确保 SliderThickness, SliderSpeed, ChkLoop, ChkPersistent, ChkObstructed 在 XAML 中已正确命名
            double thickness = (double)SliderThickness.Value;
            double speed = (double)SliderSpeed.Value;
            bool isLoop = ChkLoop.IsChecked == true;
            bool isPersistent = ChkPersistent.IsChecked == true;

            // ✨ 获取新增的“是否有遮挡”状态
            bool isObstructed = ChkObstructed.IsChecked == true;

            // 4. 初始化取消令牌
            _cts = new CancellationTokenSource();

            try
            {
                // 5. 改变 UI 状态（可选：比如禁用播放按钮，启用停止按钮）
                BtnPlay.IsEnabled = false;
                BtnStop.IsEnabled = true;

                // 6. 调用 GeometryEngine 中的双模式异步播放逻辑
                await GeometryEngine.PlaySequenceAsync(
                    _pathList.ToList(),
                    thickness,
                    speed,
                    isLoop,
                    isPersistent,
                    isObstructed,
                    _cts.Token
                );
            }
            catch (OperationCanceledException)
            {
                // 捕获用户点击“停止”引发的取消异常，无需处理
            }
            catch (System.Exception ex)
            {
                // 其他潜在异常处理
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage($"\n[回放错误]: {ex.Message}");
            }
            finally
            {
                // 7. 恢复按钮状态
                BtnPlay.IsEnabled = true;
                BtnStop.IsEnabled = false;

                if (_cts != null)
                {
                    _cts.Dispose();
                    _cts = null;
                }
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
    /// <summary>
    /// 序号转换器：将列表索引转为从 1 开始的显示序号
    /// </summary>
    public class IndexConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if (value is int index) return index + 1;
            return 1;
        }
        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
            => throw new NotImplementedException();
    }

    /// <summary>
    /// 颜色转换器：将 CAD 颜色或 Media.Color 转为 WPF 笔刷
    /// </summary>
    public class ColorToBrushConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if (value is System.Windows.Media.Color color)
                return new System.Windows.Media.SolidColorBrush(color);
            return System.Windows.Media.Brushes.Gray;
        }
        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
            => throw new NotImplementedException();
    }

    #endregion
}