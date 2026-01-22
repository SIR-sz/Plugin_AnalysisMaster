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
                PromptSelectionOptions pso = new PromptSelectionOptions();
                pso.MessageForAdding = "\n请框选或点选带有动画数据的分析线: ";

                var res = ed.GetSelection(pso);
                if (res.Status != PromptStatus.OK) return;

                using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
                {
                    // 记录本次操作已处理的指纹，避免框选内重复
                    HashSet<string> processedFingerprints = new HashSet<string>();

                    foreach (SelectedObject selObj in res.Value)
                    {
                        Entity ent = tr.GetObject(selObj.ObjectId, OpenMode.ForRead) as Entity;
                        if (ent == null || ent.ExtensionDictionary.IsNull) continue;

                        DBDictionary dict = tr.GetObject(ent.ExtensionDictionary, OpenMode.ForRead) as DBDictionary;
                        if (!dict.Contains("ANALYSIS_ANIM_DATA")) continue;

                        // ✨ 核心修改：读取该实体的唯一指纹
                        Xrecord xRec = tr.GetObject(dict.GetAt("ANALYSIS_ANIM_DATA"), OpenMode.ForRead) as Xrecord;
                        string fingerprint = "";
                        using (ResultBuffer rb = xRec.Data)
                        {
                            TypedValue[] dataArr = rb.AsArray();
                            fingerprint = dataArr[0].Value.ToString();
                        }

                        if (string.IsNullOrEmpty(fingerprint)) continue;

                        // 1. 检查本次批量选择中是否已经添加过这条线
                        if (processedFingerprints.Contains(fingerprint)) continue;

                        // 2. 检查 UI 列表中是否已经存在具有相同指纹的路径（针对多次点击的情况）
                        // 注意：这里需要我们在 CreatePathItemFromEntity 中把 Fingerprint 存入模型，或者通过 ID 反查
                        if (IsFingerprintAlreadyExists(tr, fingerprint)) continue;

                        processedFingerprints.Add(fingerprint);

                        // 解析数据构建模型
                        var item = CreatePathItemFromEntity(tr, ent, fingerprint);
                        if (item != null) _pathList.Add(item);
                    }
                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n[错误] 添加路径失败: {ex.Message}");
            }
            finally
            {
                this.Visibility = System.Windows.Visibility.Visible;
            }
        }

        // 辅助方法：检查列表中是否已存在该指纹
        private bool IsFingerprintAlreadyExists(Transaction tr, string fingerprint)
        {
            foreach (var item in _pathList)
            {
                // 获取列表中已有实体的指纹进行比对
                string existingFingerprint = GetAnimFingerprint(tr, item.Id);
                if (existingFingerprint == fingerprint) return true;
            }
            return false;
        }

        // 辅助方法：获取实体的动画指纹
        private string GetAnimFingerprint(Transaction tr, ObjectId id)
        {
            Entity ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
            if (ent == null || ent.ExtensionDictionary.IsNull) return "";

            DBDictionary dict = tr.GetObject(ent.ExtensionDictionary, OpenMode.ForRead) as DBDictionary;
            if (!dict.Contains("ANALYSIS_ANIM_DATA")) return "";

            Xrecord xRec = tr.GetObject(dict.GetAt("ANALYSIS_ANIM_DATA"), OpenMode.ForRead) as Xrecord;
            using (ResultBuffer rb = xRec.Data)
            {
                return rb.AsArray()[0].Value.ToString();
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
            if (_pathList.Count == 0)
            {
                MessageBox.Show("请先添加路径到列表。");
                return;
            }

            // 更新 UI 状态
            BtnPlay.IsEnabled = false;
            BtnStop.IsEnabled = true;
            _cts = new CancellationTokenSource();

            try
            {
                double thickness = ThicknessSlider.Value;
                bool isLoop = LoopCheck.IsChecked == true;
                bool isPersistent = PersistenceCheck.IsChecked == true;

                // 核心：调用异步引擎
                await GeometryEngine.PlaySequenceAsync(
                    _pathList.ToList(),
                    thickness,
                    isLoop,
                    isPersistent,
                    _cts.Token);
            }
            catch (OperationCanceledException)
            {
                // 正常取消，不处理
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("播放过程出错: " + ex.Message);
            }
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

        private void PathListView_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (PathListView.SelectedItem is AnimPathItem item)
            {
                ShowNumberTip(item.Id, _pathList.IndexOf(item) + 1);
            }
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