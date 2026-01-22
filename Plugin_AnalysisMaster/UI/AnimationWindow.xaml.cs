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
                    // 用于过滤本次选择中重复的组或物体
                    HashSet<string> processedInThisBatch = new HashSet<string>();
                    int addedCount = 0;

                    // 修改位置：AnimationWindow.xaml.cs -> AddPath_Click 内部循环 (约 76 行)
                    foreach (SelectedObject selObj in res.Value)
                    {
                        ObjectId actualDataHolder;
                        // ✨ 调用新版方法，获取 actualDataHolder
                        string fingerprint = GetAnimFingerprint(tr, selObj.ObjectId, out actualDataHolder);

                        if (string.IsNullOrEmpty(fingerprint)) continue;

                        if (processedInThisBatch.Contains(fingerprint) || IsFingerprintAlreadyExists(tr, fingerprint))
                            continue;

                        processedInThisBatch.Add(fingerprint);

                        // ✨ 关键修复：使用持有数据的实体来创建列表项
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

        // 辅助方法：检查列表中是否已存在该指纹
        // 修改位置：AnimationWindow.xaml.cs 约第 101 行
        private bool IsFingerprintAlreadyExists(Transaction tr, string fingerprint)
        {
            foreach (var item in _pathList)
            {
                // ✨ 修复 CS7036 错误：
                // 因为 GetAnimFingerprint 现在需要 3 个参数，这里使用 "out _" 来忽略掉不需要的 dataHolderId
                string existingFingerprint = GetAnimFingerprint(tr, item.Id, out _);

                if (existingFingerprint == fingerprint) return true;
            }
            return false;
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

        private AnimPathItem CreatePathItemFromEntity(Transaction tr, Entity ent, string fingerprint)
        {
            string layerName = ent.Layer;

            // 从指纹解析颜色
            string[] parts = fingerprint.Split('|');
            var colorStr = parts[4];
            var mediaColor = (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(colorStr);

            return new AnimPathItem
            {
                Id = ent.ObjectId,
                Name = $"图层: {layerName}",
                PathColor = mediaColor,
                GroupNumber = 1
            };
        }

        #endregion

        #region 2. 播放控制

        private async void BtnPlay_Click(object sender, RoutedEventArgs e)
        {
            if (_pathList.Count == 0) return;

            BtnPlay.IsEnabled = false;
            BtnStop.IsEnabled = true;
            _cts = new CancellationTokenSource();

            try
            {
                double thickness = ThicknessSlider.Value;
                int speed = (int)SpeedSlider.Value; // 获取倍速 (1-20)
                bool isLoop = LoopCheck.IsChecked == true;
                bool isPersistent = PersistenceCheck.IsChecked == true;

                await GeometryEngine.PlaySequenceAsync(
                    _pathList.ToList(),
                    thickness,
                    speed, // 传入倍速
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
        private void PathListView_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // A. 取消之前的高亮
            if (!_lastSelectedId.IsNull)
            {
                GeometryEngine.HighlightPath(_lastSelectedId, false);
            }

            if (PathListView.SelectedItem is AnimPathItem item)
            {
                // B. 高亮当前选中项
                GeometryEngine.HighlightPath(item.Id, true);
                _lastSelectedId = item.Id;
            }
            else
            {
                _lastSelectedId = ObjectId.Null;
            }
        }
        // 3. 确保窗口关闭时清理标签
        protected override void OnClosed(EventArgs e)
        {
            GeometryEngine.ClearPathLabels();
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