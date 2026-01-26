using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.GraphicsInterface;
using Autodesk.AutoCAD.Runtime;
using Plugin_AnalysisMaster.Models;
using System;
using System.Collections.Generic; // ✨ 新增
using System.IO;
using System.Linq;                // ✨ 新增
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Controls;
using Polyline = Autodesk.AutoCAD.DatabaseServices.Polyline;

namespace Plugin_AnalysisMaster.Services
{
    public static class GeometryEngine
    {
        // --- 新增静态变量用于管理高亮瞬态对象 ---
        private static Polyline _currentHighlightLine = null;
        private static IntegerCollection _highlightVps = new IntegerCollection();
        // ✨ 顶部增加一个静态列表，用于管理序号标签
        private static List<DBObject> _labelTransients = new List<DBObject>();
        private static IntegerCollection _labelVps = new IntegerCollection();

        private const string RegAppName = "ANALYSIS_MASTER_STYLE";
        // --- 在 GeometryEngine.cs 类中新增以下方法 ---

        private const string AnimSequenceKey = "ANALYSIS_ANIM_SEQUENCE";

        /// <summary>
        /// 将动画序列保存到 DWG 的 NOD 字典中。
        /// 改进点：步长改为 4，增加 Name (路径描述) 的持久化。
        /// </summary>
        public static void SaveSequenceToDwg(IEnumerable<AnimPathItem> items)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null || items == null) return;

            using (doc.LockDocument())
            using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
            {
                DBDictionary nod = (DBDictionary)tr.GetObject(doc.Database.NamedObjectsDictionaryId, OpenMode.ForWrite);

                Xrecord xRec = new Xrecord();
                ResultBuffer rb = new ResultBuffer();

                foreach (var item in items)
                {
                    // 顺序：1.句柄 2.组号 3.线型 4.路径描述
                    rb.Add(new TypedValue((int)DxfCode.Text, item.Id.Handle.ToString()));
                    rb.Add(new TypedValue((int)DxfCode.Int32, item.GroupNumber));
                    rb.Add(new TypedValue((int)DxfCode.Int32, (int)item.LineStyle));
                    rb.Add(new TypedValue((int)DxfCode.Text, item.Name ?? "")); // ✨ 新增：保存自定义描述
                }

                xRec.Data = rb;
                nod.SetAt(AnimSequenceKey, xRec);
                tr.AddNewlyCreatedDBObject(xRec, true);
                tr.Commit();
            }
        }

        /// <summary>
        /// 从 DWG 中读取持久化的动画序列。
        /// 改进点：步长改为 4，增加对 Name 字段的读取，并兼容旧的 3 字段数据。
        /// </summary>
        public static List<AnimPathItem> LoadSequenceFromDwg()
        {
            List<AnimPathItem> items = new List<AnimPathItem>();
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return items;

            using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
            {
                DBDictionary nod = (DBDictionary)tr.GetObject(doc.Database.NamedObjectsDictionaryId, OpenMode.ForRead);
                if (!nod.Contains(AnimSequenceKey)) return items;

                Xrecord xRec = (Xrecord)tr.GetObject(nod.GetAt(AnimSequenceKey), OpenMode.ForRead);
                using (ResultBuffer rb = xRec.Data)
                {
                    if (rb == null) return items;
                    TypedValue[] arr = rb.AsArray();

                    // ✨ 步长改为 4
                    for (int i = 0; i < arr.Length; i += 4)
                    {
                        if (i + 1 >= arr.Length) break;

                        string handleStr = arr[i].Value.ToString();
                        int groupNum = (int)arr[i + 1].Value;

                        // 线型
                        AnimLineStyle lineStyle = AnimLineStyle.Solid;
                        if (i + 2 < arr.Length) lineStyle = (AnimLineStyle)(int)arr[i + 2].Value;

                        // ✨ 路径描述：读取第 4 个字段
                        string customName = "";
                        if (i + 3 < arr.Length) customName = arr[i + 3].Value.ToString();

                        if (doc.Database.TryGetObjectId(new Handle(Convert.ToInt64(handleStr, 16)), out ObjectId id))
                        {
                            if (id.IsValid && !id.IsErased)
                            {
                                items.Add(new AnimPathItem
                                {
                                    Id = id,
                                    GroupNumber = groupNum,
                                    LineStyle = lineStyle,
                                    Name = customName // ✨ 还原描述文字
                                });
                            }
                        }
                    }
                }
                tr.Commit();
            }
            return items;
        }
        /// <summary>
        /// 专门为动画序列生成图例。
        /// 改进点：
        /// 1. 解析每一条路径关联实体的“指纹”，还原其真实的样式（阵列图块、颜色、箭头等）。
        /// 2. 调用 RenderPath 核心引擎进行渲染，确保图例示意线与图面效果完全一致。
        /// 3. 使用用户在动画列表中自定义的 Name 属性作为图例描述文字。
        /// </summary>
        public static void GenerateAnimLegend(List<AnimPathItem> animItems)
        {
            if (animItems == null || animItems.Count == 0) return;

            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            PromptPointResult ppr = ed.GetPoint("\n请指定动画图例放置起点: ");
            if (ppr.Status != PromptStatus.OK) return;

            using (DocumentLock dl = doc.LockDocument())
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

                double rowHeight = 15.0;
                double lineLength = 40.0;
                Point3d basePt = ppr.Value;
                ObjectId dashLtId = EnsureLinetypeLoaded(doc, "ANALYSIS_DASH");

                for (int i = 0; i < animItems.Count; i++)
                {
                    var item = animItems[i];

                    // 1. 获取并解析实体的样式指纹
                    Entity ent = tr.GetObject(item.Id, OpenMode.ForRead) as Entity;
                    if (ent == null || ent.ExtensionDictionary.IsNull) continue;

                    DBDictionary dict = tr.GetObject(ent.ExtensionDictionary, OpenMode.ForRead) as DBDictionary;
                    if (!dict.Contains("ANALYSIS_ANIM_DATA")) continue;

                    Xrecord xRec = tr.GetObject(dict.GetAt("ANALYSIS_ANIM_DATA"), OpenMode.ForRead) as Xrecord;
                    AnalysisStyle style = null;
                    using (ResultBuffer rb = xRec.Data)
                    {
                        if (rb == null) continue;
                        string fingerprint = rb.AsArray()[0].Value.ToString();
                        string[] parts = fingerprint.Split('|');
                        if (parts.Length < 8) continue;

                        style = new AnalysisStyle();
                        style.PathType = (PathCategory)Enum.Parse(typeof(PathCategory), parts[0]);
                        style.IsCurved = bool.Parse(parts[1]);
                        style.SelectedBlockName = parts[2];
                        style.SelectedBlockName2 = parts[3];
                        style.IsComposite = bool.Parse(parts[4]);
                        style.MainColor = (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(parts[5]);
                        style.StartArrowType = parts[6];
                        style.EndArrowType = parts[7];

                        style.ArrowSize = 1.0;
                        style.PatternScale = 1.0;
                        style.CapIndent = 2.0;
                        style.TargetLayer = ent.Layer;
                    }

                    // 2. 绘制示意图样（利用 RenderPath 还原真实样式）
                    Point3d start = new Point3d(basePt.X, basePt.Y - (i * rowHeight), 0);
                    Point3d end = new Point3d(basePt.X + lineLength, start.Y, 0);

                    using (Line samplePath = new Line(start, end))
                    {
                        ObjectIdCollection ids = new ObjectIdCollection();
                        RenderPath(btr, tr, ed, samplePath, style, ids);

                        // 如果该路径设置为虚线动画，则将生成的示意线条也设为虚线
                        if (item.LineStyle == AnimLineStyle.Dash && style.PathType == PathCategory.Solid)
                        {
                            foreach (ObjectId id in ids)
                            {
                                Entity legendEnt = tr.GetObject(id, OpenMode.ForWrite) as Entity;
                                if (legendEnt is Polyline pl)
                                {
                                    pl.LinetypeId = dashLtId;
                                    pl.LinetypeScale = 1.0;
                                }
                            }
                        }
                    }

                    // 3. 绘制描述文字（优先使用用户自定义的 item.Name）
                    MText mt = new MText();
                    mt.Contents = string.IsNullOrEmpty(item.Name) ? $"分析路径 {i + 1}" : item.Name;
                    mt.Location = new Point3d(basePt.X + lineLength + 3, start.Y, 0);
                    mt.Height = 3.0;
                    mt.Attachment = AttachmentPoint.MiddleLeft;
                    mt.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 7);

                    btr.AppendEntity(mt);
                    tr.AddNewlyCreatedDBObject(mt, true);
                }
                tr.Commit();
            }
            ed.WriteMessage($"\n[成功] 动画图例已生成。");
        }
        /// <summary>
        /// 更新图面上的路径编号标签
        /// </summary>
        public static void UpdatePathLabels(List<AnimPathItem> items, bool show)
        {
            ClearPathLabels(); // 先清理旧标签
            if (!show || items == null || items.Count == 0) return;

            Document doc = Application.DocumentManager.MdiActiveDocument;
            // 使用当前文档的事务
            using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
            {
                for (int i = 0; i < items.Count; i++)
                {
                    var item = items[i];
                    // 获取路径的采样点数据（用于定位起点）
                    var data = FetchPathData(tr, item.Id);
                    if (data == null || data.Points.Count == 0) continue;

                    Point3d startPt = data.Points[0];

                    // ✨ 调整参数：让圆圈更大，文字更清晰
                    double radius = 8.0;   // 圆圈半径增大
                    double textHeight = 10.0; // 文字高度

                    // 1. 创建圆圈
                    Circle circle = new Circle(startPt, Vector3d.ZAxis, radius);
                    circle.Color = Color.FromColorIndex(ColorMethod.ByAci, 2); // 黄色
                    circle.Thickness = 0.5; // 增加一点厚度感

                    // 2. 创建序号文字 (改用 MText 解决对齐问题)
                    MText mtext = new MText();
                    mtext.Contents = (i + 1).ToString();
                    mtext.TextHeight = textHeight;
                    mtext.Location = startPt;
                    mtext.Attachment = AttachmentPoint.MiddleCenter; // 居中对齐
                    mtext.Color = Color.FromColorIndex(ColorMethod.ByAci, 2); // 黄色

                    // 3. 添加到瞬态管理器 (Topmost 确保不被遮挡)
                    TransientManager.CurrentTransientManager.AddTransient(circle, TransientDrawingMode.DirectTopmost, 128, _labelVps);
                    TransientManager.CurrentTransientManager.AddTransient(mtext, TransientDrawingMode.DirectTopmost, 128, _labelVps);

                    _labelTransients.Add(circle);
                    _labelTransients.Add(mtext);
                }
                tr.Commit();
            }
        }
        public static void ClearPathLabels()
        {
            if (_labelTransients.Count == 0) return;
            foreach (var obj in _labelTransients)
            {
                TransientManager.CurrentTransientManager.EraseTransient(obj, _labelVps);
                obj.Dispose();
            }
            _labelTransients.Clear();
        }
        /// <summary>
        /// 使用瞬态骨架线代替实体高亮。
        /// 修改说明：
        /// <summary>
        /// 在 CAD 视图中高亮显示选中的路径。
        /// 改进点：在访问数据库前增加了 IsErased 校验，防止因操作已删除的实体引发异常。
        /// </summary>
        public static void HighlightPath(ObjectId id, bool highlight)
        {
            ClearHighlightTransient();

            // ✨ 核心修复：如果 ID 无效或已被删除，直接清理高亮并退出
            if (!highlight || id.IsNull || !id.IsValid || id.IsErased) return;

            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            Editor ed = doc.Editor;

            using (doc.LockDocument())
            using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
            {
                var data = FetchPathData(tr, id);
                if (data == null || data.Points.Count < 2) return;

                _currentHighlightLine = new Polyline();
                for (int i = 0; i < data.Points.Count; i++)
                {
                    _currentHighlightLine.AddVertexAt(i, new Point2d(data.Points[i].X, data.Points[i].Y), 0, 0, 0);
                }

                _currentHighlightLine.ColorIndex = 2; // 黄色
                _currentHighlightLine.ConstantWidth = 0.8;

                TransientManager.CurrentTransientManager.AddTransient(_currentHighlightLine, TransientDrawingMode.DirectTopmost, 128, _highlightVps);
                tr.Commit();
            }

            Autodesk.AutoCAD.Internal.Utils.FlushGraphics();
            ed.UpdateScreen();
        }
        /// <summary>
        /// 清理当前屏幕上的高亮瞬态线条。
        /// 作用：在切换路径选择或关闭窗口时，确保内存中的临时图形被正确擦除。
        /// </summary>
        public static void ClearHighlightTransient()
        {
            if (_currentHighlightLine != null)
            {
                TransientManager.CurrentTransientManager.EraseTransient(_currentHighlightLine, _highlightVps);
                _currentHighlightLine.Dispose();
                _currentHighlightLine = null;
            }
        }
        private class PathPlaybackData
        {
            public List<Point3d> Points { get; set; }
            public Autodesk.AutoCAD.Colors.Color Color { get; set; }
            public string Layer { get; set; }
            public ObjectId OriginalId { get; set; }
            // ✨ 简化：删除了 SortedMembers 和 LastShownMemberIndex
        }

        /// <summary>
        /// 异步播放序列主逻辑（智能双模式版）。
        /// 改进点：
        /// 1. 引入 isObstructed 参数：
        ///    - False（默认）：走原始逻辑，每组结束立即交接并刷新，无闪烁。
        ///    - True：走迹线常驻逻辑，播放中途静默，结束时统一交接并执行 Regen。
        /// </summary>
        public static async Task PlaySequenceAsync(
              List<AnimPathItem> items,
              double thickness,
              double speedMultiplier,
              bool isLoop,
              bool isPersistent,
              bool isObstructed, // ✨ 新参数
              CancellationToken token)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            Editor ed = doc.Editor;

            List<PathPlaybackData> allData = new List<PathPlaybackData>();
            // 预加载逻辑 (已优化)
            foreach (var item in items)
            {
                if (token.IsCancellationRequested) return;
                PathPlaybackData data = null;
                using (doc.LockDocument())
                using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
                {
                    data = FetchPathData(tr, item.Id);
                    tr.Commit();
                }
                if (data != null) allData.Add(data);
                await Task.Yield();
            }

            if (allData.Count == 0) return;

            ObjectId dashLtId = EnsureLinetypeLoaded(doc, "ANALYSIS_DASH");
            List<Polyline> activeTransients = new List<Polyline>();
            IntegerCollection vps = new IntegerCollection();

            try
            {
                do
                {
                    // 每一轮播放开始前：隐藏实体，清理旧瞬态
                    ToggleEntitiesVisibility(items.Select(i => i.Id), false);
                    ClearTransients(activeTransients, vps);

                    var groupedData = items
                        .Select(x => new { item = x, data = allData.FirstOrDefault(d => d.OriginalId == x.Id) })
                        .Where(x => x.data != null)
                        .GroupBy(x => x.item.GroupNumber)
                        .OrderBy(g => g.Key);

                    foreach (var group in groupedData)
                    {
                        if (token.IsCancellationRequested) break;

                        List<Polyline> currentGroupLines = new List<Polyline>();
                        var groupPaths = group.ToList();
                        int[] lastAddedIndices = new int[groupPaths.Count];

                        // A. 创建本组瞬态
                        foreach (var path in groupPaths)
                        {
                            Polyline pl = new Polyline { Plinegen = true };
                            pl.Color = path.data.Color;
                            pl.Layer = path.data.Layer;
                            if (path.item.LineStyle == AnimLineStyle.Dash)
                            {
                                pl.LinetypeId = dashLtId;
                                pl.LinetypeScale = 2.0;
                            }
                            pl.AddVertexAt(0, new Point2d(path.data.Points[0].X, path.data.Points[0].Y), 0, thickness, thickness);
                            pl.ConstantWidth = thickness;
                            TransientManager.CurrentTransientManager.AddTransient(pl, TransientDrawingMode.DirectTopmost, 128, vps);
                            currentGroupLines.Add(pl);
                            activeTransients.Add(pl);
                        }

                        // B. 驱动生长过程
                        for (double currentStep = speedMultiplier; ; currentStep += speedMultiplier)
                        {
                            if (token.IsCancellationRequested) break;
                            bool allFinished = true;
                            for (int i = 0; i < groupPaths.Count; i++)
                            {
                                var data = groupPaths[i].data;
                                var pl = currentGroupLines[i];
                                if (lastAddedIndices[i] >= data.Points.Count - 1) continue;
                                int nextPtIdx = Math.Min((int)currentStep, data.Points.Count - 1);
                                if (nextPtIdx > lastAddedIndices[i])
                                {
                                    int vtxIdx = pl.NumberOfVertices;
                                    pl.AddVertexAt(vtxIdx, new Point2d(data.Points[nextPtIdx].X, data.Points[nextPtIdx].Y), 0, thickness, thickness);
                                    lastAddedIndices[i] = nextPtIdx;
                                    TransientManager.CurrentTransientManager.UpdateTransient(pl, vps);
                                }
                                if (lastAddedIndices[i] < data.Points.Count - 1) allFinished = false;
                            }
                            Autodesk.AutoCAD.Internal.Utils.FlushGraphics();
                            await Task.Delay(10, token);
                            if (allFinished) break;
                        }

                        // ✨ C. 组交接逻辑（核心分支）
                        if (!isObstructed)
                        {
                            // 模式 A：无遮挡/流畅模式 -> 播完一组立即删除瞬态并恢复迹线
                            foreach (var pl in currentGroupLines)
                            {
                                TransientManager.CurrentTransientManager.EraseTransient(pl, vps);
                                activeTransients.Remove(pl);
                                pl.Dispose();
                            }
                            ToggleEntitiesVisibility(groupPaths.Select(p => p.item.Id), isPersistent);
                            await Task.Delay(20, token); // 给 UI 极短的响应时间
                        }
                        else
                        {
                            // 模式 B：有遮挡模式 -> 保持瞬态线留在屏幕最顶层，不进行实体交接
                            // 这里什么都不做，继续执行下一组即可
                        }
                    }

                    if (!isLoop) break;
                    await Task.Delay(500, token);
                } while (isLoop && !token.IsCancellationRequested);
            }
            finally
            {
                // ✨ D. 最终清理与刷新
                ClearTransients(activeTransients, vps);
                ToggleEntitiesVisibility(items.Select(i => i.Id), isPersistent);

                if (token.IsCancellationRequested || !isLoop)
                {
                    // 如果是有遮挡模式，或者用户手动停止，执行强力刷新
                    if (isObstructed)
                    {
                        ed.Regen();
                    }
                    else
                    {
                        ed.UpdateScreen();
                    }
                }
            }
        }

        /// <summary>
        /// 从实体的扩展字典中提取动画播放所需的原始数据。
        /// 改进点：增加了对象存在性校验，防止对已删除的对象（Ghost Items）执行操作。
        /// </summary>
        private static PathPlaybackData FetchPathData(Transaction tr, ObjectId id)
        {
            // ✨ 核心修复：增加 IsErased 检查
            if (id.IsNull || !id.IsValid || id.IsErased) return null;

            Entity ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
            if (ent == null || ent.ExtensionDictionary.IsNull) return null;

            DBDictionary dict = tr.GetObject(ent.ExtensionDictionary, OpenMode.ForRead) as DBDictionary;
            if (!dict.Contains("ANALYSIS_ANIM_DATA")) return null;

            Xrecord xRec = tr.GetObject(dict.GetAt("ANALYSIS_ANIM_DATA"), OpenMode.ForRead) as Xrecord;
            using (ResultBuffer rb = xRec.Data)
            {
                if (rb == null) return null;
                TypedValue[] arr = rb.AsArray();

                if (arr.Length < 2) return null;

                string fingerprint = arr[0].Value.ToString();
                string[] parts = fingerprint.Split('|');
                if (parts.Length < 6) return null;

                List<Point3d> pts = new List<Point3d>();
                for (int i = 2; i < arr.Length; i++)
                {
                    if (arr[i].TypeCode == (int)DxfCode.XCoordinate)
                    {
                        pts.Add((Point3d)arr[i].Value);
                    }
                }

                if (pts.Count < 2) return null;

                try
                {
                    var colorStr = parts[5];
                    var mediaColor = (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(colorStr);

                    return new PathPlaybackData
                    {
                        OriginalId = id,
                        Points = pts,
                        Color = Autodesk.AutoCAD.Colors.Color.FromRgb(mediaColor.R, mediaColor.G, mediaColor.B),
                        Layer = ent.Layer
                    };
                }
                catch
                {
                    return new PathPlaybackData
                    {
                        OriginalId = id,
                        Points = pts,
                        Color = ent.Color,
                        Layer = ent.Layer
                    };
                }
            }
        }
        /// <summary>
        /// 将指纹字符串解析为 Point3d 坐标列表。
        /// 作用：处理形如 "PathType|x,y,z;x,y,z;...|Style" 的长字符串。
        /// 它是所有功能（播放、编号、高亮）获取几何数据的唯一入口。
        /// </summary>
        private static List<Point3d> GetFingerprintPoints(string fingerprint)
        {
            List<Point3d> points = new List<Point3d>();
            if (string.IsNullOrEmpty(fingerprint)) return points;

            try
            {
                string[] parts = fingerprint.Split('|');
                // 指纹至少应包含类型和坐标段
                if (parts.Length < 2) return points;

                // 坐标段在 index 为 1 的位置
                string[] pts = parts[1].Split(';');
                foreach (string p in pts)
                {
                    if (string.IsNullOrWhiteSpace(p)) continue;

                    string[] c = p.Split(',');
                    if (c.Length == 3)
                    {
                        if (double.TryParse(c[0], out double x) &&
                            double.TryParse(c[1], out double y) &&
                            double.TryParse(c[2], out double z))
                        {
                            points.Add(new Point3d(x, y, z));
                        }
                    }
                }
            }
            catch
            {
                // 异常时返回已解析部分
            }
            return points;
        }
        // 辅助方法：清理瞬态
        private static void ClearTransients(List<Polyline> lines, IntegerCollection vps)
        {
            foreach (var line in lines)
            {
                TransientManager.CurrentTransientManager.EraseTransient(line, vps);
                line.Dispose();
            }
        }

        /// <summary>
        /// 批量切换实体的可见性。
        /// 改进点：移除内部所有强力刷新指令，使其在循环播放过程中保持静默，避免闪烁。
        /// </summary>
        private static void ToggleEntitiesVisibility(IEnumerable<ObjectId> ids, bool visible)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            using (doc.LockDocument())
            using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
            {
                HashSet<ObjectId> allToToggle = new HashSet<ObjectId>();
                foreach (var id in ids)
                {
                    if (id.IsNull || !id.IsValid || id.IsErased) continue;
                    allToToggle.Add(id);
                    Entity ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                    if (ent != null)
                    {
                        ObjectIdCollection reactors = ent.GetPersistentReactorIds();
                        if (reactors != null)
                        {
                            foreach (ObjectId rId in reactors)
                            {
                                if (rId.IsValid && !rId.IsErased && rId.ObjectClass.IsDerivedFrom(RXClass.GetClass(typeof(Group))))
                                {
                                    Group gp = tr.GetObject(rId, OpenMode.ForRead) as Group;
                                    if (gp != null) foreach (ObjectId mId in gp.GetAllEntityIds()) allToToggle.Add(mId);
                                }
                            }
                        }
                    }
                }

                foreach (var tId in allToToggle)
                {
                    try
                    {
                        Entity target = tr.GetObject(tId, OpenMode.ForWrite) as Entity;
                        if (target != null) target.Visible = visible;
                    }
                    catch { continue; }
                }
                tr.Commit();
            }
            // 仅执行显存刷新，不触发重生成
            Autodesk.AutoCAD.Internal.Utils.FlushGraphics();
        }

        /// <summary>
        /// 动线绘制主入口：支持多点连续拾取与自动打组。
        /// 修改点：在事务提交前增加了绘图顺序置顶逻辑，确保绘制完后线条立即浮在图片背景上方。
        /// </summary>
        public static void DrawAnalysisLine(Point3dCollection inputPoints, AnalysisStyle style)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Point3dCollection points = inputPoints;

            if (points == null)
            {
                PromptPointOptions ppo = new PromptPointOptions("\n指定起点: ");
                PromptPointResult ppr = ed.GetPoint(ppo);
                if (ppr.Status != PromptStatus.OK) return;

                var jig = new AnalysisLineJig(ppr.Value, style);
                while (true)
                {
                    var res = ed.Drag(jig);
                    if (res.Status == PromptStatus.OK)
                    {
                        jig.AddPoint(jig.LastPoint);
                    }
                    else if (res.Status == PromptStatus.None)
                    {
                        jig.AddPoint(jig.LastPoint);
                        break;
                    }
                    else break;
                }
                points = jig.GetPoints();
            }

            if (points == null || points.Count < 2) return;

            using (DocumentLock dl = doc.LockDocument())
            {
                Database db = doc.Database;
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    EnsureLayer(db, tr, style);
                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                    ObjectIdCollection allIds = new ObjectIdCollection();

                    Curve path = CreatePathCurve(points, style.IsCurved);
                    using (path)
                    {
                        RenderPath(btr, tr, ed, path, style, allIds);
                        if (allIds.Count > 0)
                        {
                            Entity headEnt = tr.GetObject(allIds[0], OpenMode.ForWrite) as Entity;
                            if (headEnt != null)
                            {
                                List<Point3d> skeletonPoints = style.IsAnimated ? SampleCurve(path, style.SamplingInterval) : null;
                                WriteFingerprint(tr, headEnt, style, skeletonPoints);
                            }
                        }
                    }

                    if (allIds.Count > 1)
                    {
                        DBDictionary gd = (DBDictionary)tr.GetObject(db.GroupDictionaryId, OpenMode.ForWrite);
                        Group grp = new Group("分析动线单元", true);
                        gd.SetAt("*", grp);
                        tr.AddNewlyCreatedDBObject(grp, true);
                        grp.Append(allIds);
                    }

                    // ✨ 核心修复：绘制完成后强制置顶
                    if (allIds.Count > 0)
                    {
                        ObjectId dotId = btr.DrawOrderTableId;
                        if (!dotId.IsNull)
                        {
                            DrawOrderTable dot = tr.GetObject(dotId, OpenMode.ForWrite) as DrawOrderTable;
                            dot.MoveToTop(allIds);
                        }
                    }

                    tr.Commit();
                }
            }
            Autodesk.AutoCAD.Internal.Utils.FlushGraphics();
        }
        /// <summary>
        /// 渲染路径封装：统一处理修剪、主体渲染和端头渲染。
        /// 修改逻辑：增加 Editor 参数接收，并将其透传给主体渲染方法。
        /// </summary>
        public static void RenderPath(BlockTableRecord btr, Transaction tr, Editor ed, Curve path, AnalysisStyle style, ObjectIdCollection allIds)
        {
            double indent = style.CapIndent;
            // 修剪路径（处理缩进）
            Curve trimmedPath = TrimPath(path, indent, indent);

            if (trimmedPath != null)
            {
                // 1. 渲染主体（实线或阵列）
                // ✨ 修改：透传 ed 参数
                RenderBody(btr, tr, ed, trimmedPath, style, allIds);

                // 2. 渲染起始端
                RenderHead(btr, tr, path, style, allIds);

                // 3. 渲染结束端
                RenderTail(btr, tr, path, style, allIds);

                // 如果生成了新曲线则释放
                if (trimmedPath != path) trimmedPath.Dispose();
            }
        }
        /// <summary>
        /// 渲染动线主体部分。
        /// 修改逻辑：在调用 AddToDb 时传入原始 curve 路径，以便执行动画采样点的存储逻辑。
        /// </summary>
        private static void RenderBody(BlockTableRecord btr, Transaction tr, Editor ed, Curve curve, AnalysisStyle style, ObjectIdCollection ids)
        {
            double len = curve.GetDistanceAtParameter(curve.EndParam);

            if (style.PathType == PathCategory.Solid)
            {
                // ✨ 检查是否为恒定宽度
                bool isConstantWidth = Math.Abs(style.StartWidth - style.MidWidth) < 0.001 &&
                                       Math.Abs(style.MidWidth - style.EndWidth) < 0.001;

                List<Point3d> samplePoints;

                if (!style.IsCurved && isConstantWidth)
                {
                    // ✨ 情况 A：折线模式且宽度恒定 -> 直接提取原始顶点（最精简）
                    using (Polyline plRaw = ConvertToPolyline(curve))
                    {
                        samplePoints = new List<Point3d>();
                        for (int i = 0; i < plRaw.NumberOfVertices; i++)
                        {
                            samplePoints.Add(plRaw.GetPoint3dAt(i));
                        }
                    }
                }
                else
                {
                    // ✨ 情况 B：曲线模式或宽度渐变 -> 使用自适应采样（确保顺滑）
                    samplePoints = SampleCurveAdaptively(curve);
                }

                using (Polyline pl = new Polyline { Plinegen = true })
                {
                    for (int i = 0; i < samplePoints.Count; i++)
                    {
                        Point3d pt = samplePoints[i];
                        double t = curve.GetDistAtPoint(pt) / len;

                        // 计算渐变宽度
                        double w = CalculateBezierWidth(t, style.StartWidth, style.MidWidth, style.EndWidth);
                        pl.AddVertexAt(i, new Point2d(pt.X, pt.Y), 0, w, w);
                    }
                    ids.Add(AddToDb(btr, tr, pl, style, null));
                }
            }
            else if (style.PathType == PathCategory.Pattern)
            {
                // 导入所需的图块资源
                ObjectId blockId1 = GetOrImportBlock(style.SelectedBlockName, btr.Database, tr);
                ObjectId blockId2 = style.IsComposite ? GetOrImportBlock(style.SelectedBlockName2, btr.Database, tr) : blockId1;

                if (blockId1.IsNull) return;

                double spacing = style.PatternSpacing;
                if (spacing <= 0.001) spacing = 10.0; // 防止间距过小导致死循环

                int count = (int)(len / spacing) + 1;
                // 计算边距使阵列居中分布
                double margin = (len - (count - 1) * spacing) / 2.0;

                for (int i = 0; i < count; i++)
                {
                    double currentDist = Math.Max(0, Math.Min(len, margin + (i * spacing)));

                    // 处理组合模式：奇数位使用备用图块
                    ObjectId currentBlockId = (i % 2 != 0 && style.IsComposite) ? blockId2 : blockId1;
                    if (currentBlockId.IsNull) currentBlockId = blockId1;

                    Point3d pt = curve.GetPointAtDist(currentDist);
                    // 计算切线方向以确定图块旋转角度
                    Vector3d tan = curve.GetFirstDerivative(curve.GetParameterAtDistance(currentDist)).GetNormal();

                    using (BlockReference br = new BlockReference(pt, currentBlockId))
                    {
                        br.Rotation = Math.Atan2(tan.Y, tan.X);
                        br.ScaleFactors = new Scale3d(style.PatternScale);

                        // ✨ 核心修复：阵列模式下的子单元传 null，避免重复采样动画点
                        ids.Add(AddToDb(btr, tr, br, style, null));
                    }
                }
            }
        }
        /// <summary>
        /// 自适应采样：彻底精简直线，仅在弯道处加密。
        /// 修改说明：移除了 maxDist 限制，防止长直段被切分。
        /// </summary>
        private static List<Point3d> SampleCurveAdaptively(Curve curve, double angleToleranceDeg = 0.5)
        {
            List<Point3d> pts = new List<Point3d>();
            double totalLen = curve.GetDistanceAtParameter(curve.EndParam);

            // 基础采样步长
            double step = Math.Min(totalLen / 100.0, 1.0);
            double toleranceCos = Math.Cos(angleToleranceDeg * Math.PI / 180.0);

            pts.Add(curve.StartPoint);
            Point3d lastSavedPt = curve.StartPoint;
            Vector3d lastSavedDir = curve.GetFirstDerivative(curve.StartParam).GetNormal();

            for (double d = step; d < totalLen; d += step)
            {
                Point3d currentPt = curve.GetPointAtDist(d);
                Vector3d currentDir = curve.GetFirstDerivative(curve.GetParameterAtDistance(d)).GetNormal();

                // ✨ 仅当方向变化超过阈值时才添加点（去掉了距离限制）
                if (currentDir.DotProduct(lastSavedDir) < toleranceCos)
                {
                    pts.Add(currentPt);
                    lastSavedPt = currentPt;
                    lastSavedDir = currentDir;
                }
            }

            if (pts.Count > 0 && pts.Last().DistanceTo(curve.EndPoint) > 0.001)
                pts.Add(curve.EndPoint);

            return pts;
        }
        /// <summary>
        /// 动画回放主入口：采用纯命令行交互，解决 UI 弹出框干扰及光标卡顿感。
        /// 修改逻辑：
        /// 1. 纯净命令行提示：移除 PromptKeywordOptions 的 Message，转而使用 ed.WriteMessage 输出引导语，避免触发动态输入框（Dynamic Input）。
        /// 2. 自由移动光标：在 GetKeywords 等待期间，允许用户自由移动鼠标将十字光标移开视口中心。
        /// 3. 增强刷新：确认后立即进入隐藏和动画流程，确保视觉上的无缝衔接。
        /// </summary>
        public static void PlayPathAnimation()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            Editor ed = doc.Editor;

            ed.WriteMessage("\n[回放调试] 请在图中拾取分析线...");

            // 1. 拾取带有动画数据的对象
            PromptEntityOptions peo = new PromptEntityOptions("\n请选择需要回放动画的分析线: ");
            peo.SetRejectMessage("\n所选对象不含动画路径数据。");
            var res = ed.GetEntity(peo);
            if (res.Status != PromptStatus.OK) return;

            // ✨ 2. 纯命令行确认逻辑：不使用弹出式消息框
            ed.WriteMessage("\n------------------------------------------------------------");
            ed.WriteMessage("\n[确认回放] 请将鼠标移开视口中心，按下 [回车] 或 [空格] 开始播放...");
            ed.WriteMessage("\n------------------------------------------------------------");

            // 使用空消息的关键词选项，从而不在光标旁显示 UI 浮窗
            PromptKeywordOptions pko = new PromptKeywordOptions("");
            pko.AllowNone = true;              // 允许直接按回车/空格
            pko.AppendKeywordsToMessage = false; // 不在命令行末尾追加关键词提示

            var pkr = ed.GetKeywords(pko);
            if (pkr.Status != PromptStatus.OK && pkr.Status != PromptStatus.None) return;

            // 3. 进入锁定和回放流程
            using (doc.LockDocument())
            {
                using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
                {
                    Entity pickedEnt = tr.GetObject(res.ObjectId, OpenMode.ForWrite) as Entity;
                    if (pickedEnt == null || pickedEnt.ExtensionDictionary.IsNull) return;

                    // 读取动画数据
                    DBDictionary extDict = tr.GetObject(pickedEnt.ExtensionDictionary, OpenMode.ForRead) as DBDictionary;
                    if (!extDict.Contains("ANALYSIS_ANIM_DATA"))
                    {
                        ed.WriteMessage("\n[错误] 未找到动画路径数据。");
                        return;
                    }

                    Xrecord xRec = tr.GetObject(extDict.GetAt("ANALYSIS_ANIM_DATA"), OpenMode.ForRead) as Xrecord;
                    using (ResultBuffer rb = xRec.Data)
                    {
                        if (rb == null) return;

                        TypedValue[] dataArr = rb.AsArray();
                        string fingerprint = dataArr[0].Value.ToString();
                        Point3dCollection pts = new Point3dCollection();
                        for (int i = 2; i < dataArr.Length; i++)
                        {
                            if (dataArr[i].TypeCode == (int)DxfCode.XCoordinate)
                                pts.Add((Point3d)dataArr[i].Value);
                        }

                        if (pts.Count < 2) return;

                        // 4. 识别并隐藏关联的所有实体
                        List<Entity> entitiesToHide = new List<Entity>();
                        entitiesToHide.Add(pickedEnt);

                        ObjectIdCollection reactorIds = pickedEnt.GetPersistentReactorIds();
                        if (reactorIds != null)
                        {
                            foreach (ObjectId id in reactorIds)
                            {
                                if (id.ObjectClass.IsDerivedFrom(RXClass.GetClass(typeof(Group))))
                                {
                                    Group gp = tr.GetObject(id, OpenMode.ForRead) as Group;
                                    if (gp != null)
                                    {
                                        foreach (ObjectId memberId in gp.GetAllEntityIds())
                                        {
                                            Entity member = tr.GetObject(memberId, OpenMode.ForWrite) as Entity;
                                            if (member != null && !entitiesToHide.Contains(member))
                                                entitiesToHide.Add(member);
                                        }
                                    }
                                }
                            }
                        }

                        // 批量隐藏
                        foreach (var ent in entitiesToHide) ent.Visible = false;
                        tr.TransactionManager.QueueForGraphicsFlush();
                        Autodesk.AutoCAD.Internal.Utils.FlushGraphics();
                        ed.UpdateScreen();

                        // 5. 执行增量瞬态生长动画
                        using (Polyline animLine = new Polyline())
                        {
                            // 使用黄色高亮
                            animLine.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 2);
                            animLine.Layer = pickedEnt.Layer;
                            animLine.AddVertexAt(0, new Point2d(pts[0].X, pts[0].Y), 0, 0, 0);

                            IntegerCollection vps = new IntegerCollection();
                            Autodesk.AutoCAD.GraphicsInterface.TransientManager.CurrentTransientManager.AddTransient(
                                animLine, Autodesk.AutoCAD.GraphicsInterface.TransientDrawingMode.DirectTopmost, 128, vps);

                            try
                            {
                                for (int i = 1; i < pts.Count; i++)
                                {
                                    animLine.AddVertexAt(i, new Point2d(pts[i].X, pts[i].Y), 0, 0, 0);
                                    Autodesk.AutoCAD.GraphicsInterface.TransientManager.CurrentTransientManager.UpdateTransient(animLine, vps);

                                    Autodesk.AutoCAD.Internal.Utils.FlushGraphics();
                                    Autodesk.AutoCAD.ApplicationServices.Application.UpdateScreen();

                                    System.Threading.Thread.Sleep(20);
                                }
                            }
                            finally
                            {
                                // 6. 清理并恢复显示
                                Autodesk.AutoCAD.GraphicsInterface.TransientManager.CurrentTransientManager.EraseTransient(animLine, vps);
                                foreach (var ent in entitiesToHide) ent.Visible = true;
                                ed.UpdateScreen();
                            }
                        }
                    }
                    tr.Commit();
                }
            }
        }
        /// <summary>
        /// 在路径起点渲染选中的端头图块。
        /// 修改逻辑：将缩放比例从 PatternScale 更改为 ArrowSize，实现缩放控制的解耦。
        /// 同时保持起始端 180 度反向逻辑，确保箭头背向路径。
        /// </summary>
        private static void RenderHead(BlockTableRecord btr, Transaction tr, Curve curve, AnalysisStyle style, ObjectIdCollection idCol)
        {
            if (string.IsNullOrEmpty(style.StartArrowType) || style.StartArrowType == "None") return;

            ObjectId blockId = GetOrImportBlock(style.StartArrowType, btr.Database, tr);
            if (blockId.IsNull) return;

            Point3d startPt = curve.StartPoint;
            Vector3d dir = curve.GetFirstDerivative(curve.StartParam).GetNormal();

            using (BlockReference br = new BlockReference(startPt, blockId))
            {
                br.Rotation = Math.Atan2(dir.Y, dir.X) + Math.PI;

                // ✨ 修改：端头缩放现在独立使用 style.ArrowSize
                br.ScaleFactors = new Scale3d(style.ArrowSize);

                br.Color = Autodesk.AutoCAD.Colors.Color.FromRgb(
                    style.MainColor.R, style.MainColor.G, style.MainColor.B);

                ObjectId id = btr.AppendEntity(br);
                tr.AddNewlyCreatedDBObject(br, true);
                idCol.Add(id);
            }
        }
        /// <summary>
        /// 在路径终点渲染选中的端头图块。
        /// 修改逻辑：将缩放比例从 PatternScale 更改为 ArrowSize，确保端头缩放不受阵列单元缩放影响。
        /// </summary>
        private static void RenderTail(BlockTableRecord btr, Transaction tr, Curve curve, AnalysisStyle style, ObjectIdCollection idCol)
        {
            if (string.IsNullOrEmpty(style.EndArrowType) || style.EndArrowType == "None") return;

            ObjectId blockId = GetOrImportBlock(style.EndArrowType, btr.Database, tr);
            if (blockId.IsNull) return;

            Point3d endPt = curve.EndPoint;
            Vector3d dir = curve.GetFirstDerivative(curve.EndParam).GetNormal();

            using (BlockReference br = new BlockReference(endPt, blockId))
            {
                br.Rotation = Math.Atan2(dir.Y, dir.X);

                // ✨ 修改：端头缩放现在独立使用 style.ArrowSize
                br.ScaleFactors = new Scale3d(style.ArrowSize);

                br.Color = Autodesk.AutoCAD.Colors.Color.FromRgb(
                    style.MainColor.R, style.MainColor.G, style.MainColor.B);

                // ✨ 新增：设置图层，确保端头和线条在同一个图层
                br.Layer = style.TargetLayer;

                ObjectId id = btr.AppendEntity(br);
                tr.AddNewlyCreatedDBObject(br, true);
                idCol.Add(id);
            }
        }
        private static void RenderFeature(BlockTableRecord btr, Transaction tr, Point3dCollection pts, AnalysisStyle style, ObjectIdCollection idCol)
        {
            if (pts.Count == 0) return;
            using (Polyline pl = new Polyline())
            {
                double bulge = (pts.Count == 2) ? 1.0 : 0.0; // 如果是 2 个点则通过凸度画圆
                for (int i = 0; i < pts.Count; i++)
                    pl.AddVertexAt(i, new Point2d(pts[i].X, pts[i].Y), bulge, 0, 0);
                pl.Closed = true;
                idCol.Add(AddToDb(btr, tr, pl, style));
            }
        }

        /// <summary>
        /// 计算特征点位：原本用于内置箭头的几何计算。
        /// 由于现在改用图块渲染，此方法在当前版本中仅作为内部辅助（或可直接删除相关调用）。
        /// 修改逻辑：移除对已不存在的样式属性的引用，防止编译错误。
        /// </summary>
        private static Point3dCollection CalculateFeaturePoints(Point3d basePt, Vector3d dir, AnalysisStyle style, bool isHead)
        {
            // ✨ 简化逻辑：因为现在使用 RenderHead/RenderTail (图块模式)，不再需要这个复杂的 Switch 计算。
            // 这里返回空集合，旧的 RenderFeature 调用将不会产生任何对象。
            return new Point3dCollection();
        }

        private static ObjectId GetOrImportBlock(string name, Database destDb, Transaction tr)
        {
            if (string.IsNullOrEmpty(name)) return ObjectId.Null;
            BlockTable bt = (BlockTable)tr.GetObject(destDb.BlockTableId, OpenMode.ForRead);
            if (bt.Has(name)) return bt[name];

            string dllPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string libPath = Path.Combine(dllPath, "Assets", "PatternLibrary.dwg");
            if (!File.Exists(libPath)) return ObjectId.Null;

            using (Database srcDb = new Database(false, true))
            {
                // 兼容 2015 版 API
                srcDb.ReadDwgFile(libPath, FileShare.Read, true, "");
                using (var trSrc = srcDb.TransactionManager.StartTransaction())
                {
                    BlockTable btSrc = (BlockTable)trSrc.GetObject(srcDb.BlockTableId, OpenMode.ForRead);
                    if (btSrc.Has(name))
                    {
                        ObjectIdCollection ids = new ObjectIdCollection { btSrc[name] };
                        IdMapping map = new IdMapping();
                        destDb.WblockCloneObjects(ids, destDb.BlockTableId, map, DuplicateRecordCloning.Ignore, false);
                        return map[btSrc[name]].Value;
                    }
                }
            }
            return ObjectId.Null;
        }

        /// <summary>
        /// 核心方法：根据绘制点位和样式，生成实体并存入数据库
        // --- 修改后的 FinalizeAndSave ---
        // 文件：GeometryEngine.cs -> FinalizeAndSave 方法

        public static void FinalizeAndSave(List<Point3d> points, AnalysisStyle style, List<Point3d> sampledPoints)
        {
            if (points == null || points.Count < 2) return;

            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;

            using (DocumentLock loc = doc.LockDocument())
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTableRecord modelSpace = (BlockTableRecord)tr.GetObject(
                    SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForWrite);

                Point3dCollection ptsCol = new Point3dCollection(points.ToArray());
                Curve pathCurve = CreatePathCurve(ptsCol, style.IsCurved);

                using (pathCurve)
                {
                    ObjectIdCollection ids = new ObjectIdCollection();

                    // 1. 渲染所有几何（此时 AddToDb 仅负责画线，不写动画数据）
                    if (style.PathType == PathCategory.Solid)
                        RenderBody(modelSpace, tr, doc.Editor, pathCurve, style, ids);
                    else
                        RenderPath(modelSpace, tr, doc.Editor, pathCurve, style, ids);

                    // 2. ✨ 核心逻辑：只在第一个实体写入动画指纹
                    // 这样做确保了：Picking 时只添加一个条目到列表，且数据来源于骨架线采样
                    if (ids.Count > 0 && style.IsAnimated && sampledPoints != null)
                    {
                        Entity headEnt = tr.GetObject(ids[0], OpenMode.ForWrite) as Entity;
                        if (headEnt != null)
                        {
                            WriteFingerprint(tr, headEnt, style, sampledPoints);
                        }
                    }

                    // 3. 自动打组（保持 CAD 操作的整体性）
                    if (ids.Count > 1)
                    {
                        Group gp = new Group("AnalysisGroup", true);
                        DBDictionary groupDict = (DBDictionary)tr.GetObject(db.GroupDictionaryId, OpenMode.ForWrite);
                        groupDict.SetAt("*", gp);
                        tr.AddNewlyCreatedDBObject(gp, true);
                        gp.Append(ids);
                    }
                }
                tr.Commit();
            }
        }
        // --- 3. 新版列表写入方法 (用于 FinalizeAndSave 内部调用) ---
        private static void SaveEntitiesToDbInternal(Transaction tr, BlockTableRecord btr, List<Entity> entities, AnalysisStyle style, List<Point3d> sampledPoints)
        {
            if (entities == null || entities.Count == 0) return;

            ObjectIdCollection ids = new ObjectIdCollection();
            foreach (Entity ent in entities)
            {
                if (ent.ObjectId.IsNull)
                {
                    btr.AppendEntity(ent);
                    tr.AddNewlyCreatedDBObject(ent, true);
                }
                ids.Add(ent.ObjectId);
            }

            // 自动打组逻辑
            if (ids.Count > 1)
            {
                Group gp = new Group("AnalysisGroup", true);
                DBDictionary groupDict = (DBDictionary)tr.GetObject(btr.Database.GroupDictionaryId, OpenMode.ForWrite);
                groupDict.SetAt("*", gp);
                tr.AddNewlyCreatedDBObject(gp, true);
                gp.Append(ids);
            }

            // 在第一个实体上写指纹
            WriteFingerprint(tr, entities[0], style, sampledPoints);
        }
        /// <summary>
        /// 优化后的 AddToDb：实现“高性能打组”与“单指纹写入”
        /// 修改点：增加了绘图顺序控制逻辑，确保阵列生成的子元素不被背景图片遮挡。
        /// </summary>
        public static void AddToDb(List<Entity> entities, AnalysisStyle style, List<Point3d> sampledPoints)
        {
            if (entities == null || entities.Count == 0) return;
            Database db = HostApplicationServices.WorkingDatabase;
            using (var tr = db.TransactionManager.StartTransaction())
            {
                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForWrite);

                ObjectIdCollection ids = new ObjectIdCollection();
                foreach (Entity ent in entities)
                {
                    if (ent.ObjectId.IsNull)
                    {
                        btr.AppendEntity(ent);
                        tr.AddNewlyCreatedDBObject(ent, true);
                    }
                    ids.Add(ent.ObjectId);
                }

                // 自动打组
                if (ids.Count > 1)
                {
                    Group gp = new Group("AnalysisGroup", true);
                    DBDictionary groupDict = (DBDictionary)tr.GetObject(db.GroupDictionaryId, OpenMode.ForWrite);
                    groupDict.SetAt("*", gp);
                    tr.AddNewlyCreatedDBObject(gp, true);
                    gp.Append(ids);
                }

                // ✨ 核心修复：批量入库时执行置顶
                if (ids.Count > 0)
                {
                    ObjectId dotId = btr.DrawOrderTableId;
                    if (!dotId.IsNull)
                    {
                        DrawOrderTable dot = tr.GetObject(dotId, OpenMode.ForWrite) as DrawOrderTable;
                        dot.MoveToTop(ids);
                    }
                }

                // 仅在第一个实体写指纹
                WriteFingerprint(tr, entities[0], style, sampledPoints);
                tr.Commit();
            }
            Autodesk.AutoCAD.Internal.Utils.FlushGraphics();
        }


        // --- 修复：支持单体保存并返回 ObjectId 的重载 ---
        // 解决错误：CS1503 (void 无法转换为 ObjectId) 和 CS1501 (没有 5 个参数的重载)
        // --- 重载 2：给单条实体用的（旧代码兼容重载） ---
        // 解决 CS1503 (void 无法转换为 ObjectId) 和 CS1501 (参数不匹配)
        // --- 修改后的 AddToDb (单体版本) ---
        // --- 请确保这里的返回类型是 ObjectId 而不是 void ---
        // 文件：GeometryEngine.cs

        public static ObjectId AddToDb(BlockTableRecord btr, Transaction tr, Entity ent, AnalysisStyle style, Curve rawPath = null)
        {
            if (ent == null) return ObjectId.Null;

            // ✨ 修复 1：统一设置颜色和图层
            ent.Color = Autodesk.AutoCAD.Colors.Color.FromRgb(style.MainColor.R, style.MainColor.G, style.MainColor.B);
            ent.Layer = style.TargetLayer;

            if (ent.ObjectId.IsNull)
            {
                btr.AppendEntity(ent);
                tr.AddNewlyCreatedDBObject(ent, true);
            }

            // ✨ 优化：如果此处传入了 rawPath（兼容老逻辑），则执行采样
            // 但在新的 FinalizeAndSave 逻辑中，我们会传 null，从而跳过这里的采样，提升速度
            List<Point3d> sampledPoints = null;
            if (style.IsAnimated && rawPath != null)
            {
                sampledPoints = SampleCurve(rawPath, style.SamplingInterval);
            }

            // 如果有采样点，则写入指纹（单线模式）
            if (sampledPoints != null)
            {
                WriteFingerprint(tr, ent, style, sampledPoints);
            }

            return ent.ObjectId;
        }

        // 辅助采样方法
        private static List<Point3d> SampleCurve(Curve curve, double interval)
        {
            List<Point3d> points = new List<Point3d>();
            try
            {
                double len = curve.GetDistanceAtParameter(curve.EndParam);
                double step = interval > 0 ? interval : 20.0;
                int count = (int)Math.Max(len / step, 10);
                for (int i = 0; i <= count; i++)
                {
                    points.Add(curve.GetPointAtDist(Math.Min(len, i * step)));
                }
            }
            catch { }
            return points;
        }
        // --- 3. 抽离出的指纹写入逻辑（私有辅助） ---
        // 文件：GeometryEngine.cs -> WriteFingerprint 方法
        private static void WriteFingerprint(Transaction tr, Entity ent, AnalysisStyle style, List<Point3d> sampledPoints)
        {
            if (ent.ExtensionDictionary.IsNull)
            {
                ent.UpgradeOpen();
                ent.CreateExtensionDictionary();
            }

            DBDictionary dict = (DBDictionary)tr.GetObject(ent.ExtensionDictionary, OpenMode.ForWrite);

            // ✨ 核心修复：直接使用 style.MainColor.ToString() 替代 ColorSerializationHelper
            // 它会生成类似 "#AARRGGBB" 的字符串，这是最通用的格式
            string fingerprint = $"{style.PathType}|{style.IsCurved}|{style.SelectedBlockName}|{style.SelectedBlockName2}|{style.IsComposite}|" +
                                 $"{style.MainColor.ToString()}|{style.StartArrowType}|{style.EndArrowType}";

            using (Xrecord xRec = new Xrecord())
            {
                ResultBuffer rb = new ResultBuffer();
                rb.Add(new TypedValue((int)DxfCode.Text, fingerprint));

                int count = (sampledPoints != null) ? sampledPoints.Count : 0;
                rb.Add(new TypedValue((int)DxfCode.Int32, count));

                if (sampledPoints != null)
                {
                    foreach (Point3d pt in sampledPoints)
                    {
                        rb.Add(new TypedValue((int)DxfCode.XCoordinate, pt));
                    }
                }

                xRec.Data = rb;
                dict.SetAt("ANALYSIS_ANIM_DATA", xRec);
                tr.AddNewlyCreatedDBObject(xRec, true);
            }
        }
        /// <summary>
        /// 全局图例生成逻辑：从扩展字典读取样式并生成图例
        /// </summary>
        public static void GenerateLegend()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            using (DocumentLock dl = doc.LockDocument())
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                // 1. 搜寻图中所有分析线样式（从扩展字典中搜寻）
                var uniqueStyles = new Dictionary<string, AnalysisStyle>();
                BlockTableRecord curSpace = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForRead);

                foreach (ObjectId id in curSpace)
                {
                    Entity ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                    if (ent == null || ent.ExtensionDictionary.IsNull) continue;

                    DBDictionary dict = tr.GetObject(ent.ExtensionDictionary, OpenMode.ForRead) as DBDictionary;
                    if (!dict.Contains("ANALYSIS_ANIM_DATA")) continue;

                    Xrecord xRec = tr.GetObject(dict.GetAt("ANALYSIS_ANIM_DATA"), OpenMode.ForRead) as Xrecord;
                    using (ResultBuffer rb = xRec.Data)
                    {
                        if (rb == null) continue;
                        TypedValue[] dataArr = rb.AsArray();
                        string fingerprint = dataArr[0].Value.ToString();

                        if (!uniqueStyles.ContainsKey(fingerprint))
                        {
                            string[] parts = fingerprint.Split('|');
                            // ✨ 确保数组长度足够（现在至少有 8 个字段）
                            if (parts.Length < 8) continue;

                            var style = new AnalysisStyle();

                            // ✨ 严格按照 WriteFingerprint 的顺序重新对齐索引：
                            // 顺序：0:PathType | 1:IsCurved | 2:Block1 | 3:Block2 | 4:IsComposite | 5:Color | 6:StartArrow | 7:EndArrow

                            style.PathType = (PathCategory)Enum.Parse(typeof(PathCategory), parts[0]);
                            style.IsCurved = bool.Parse(parts[1]);      // 索引 1 是布尔值
                            style.SelectedBlockName = parts[2];         // 索引 2 是字符串
                            style.SelectedBlockName2 = parts[3];        // 索引 3 是字符串
                            style.IsComposite = bool.Parse(parts[4]);   // 索引 4 是布尔值

                            var colorStr = parts[5];                    // 索引 5 是颜色字符串
                            style.MainColor = (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(colorStr);

                            style.StartArrowType = parts[6];            // 索引 6 是字符串
                            style.EndArrowType = parts[7];              // 索引 7 是字符串

                            // 固定图例显示参数
                            style.ArrowSize = 1.0;
                            style.PatternScale = 1.0;
                            style.CapIndent = 2.0;
                            style.IsAnimated = false;

                            uniqueStyles.Add(fingerprint, style);
                        }
                    }
                }

                if (uniqueStyles.Count == 0)
                {
                    ed.WriteMessage("\n[提示] 图中未检测到带有动画数据的分析线，无法生成图例。");
                    return;
                }

                // 2. 指定放置位置
                PromptPointResult ppr = ed.GetPoint("\n请指定图例放置起点: ");
                if (ppr.Status != PromptStatus.OK) return;

                Point3d insertPt = ppr.Value;
                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

                // 3. 绘制图例
                double rowHeight = 15.0; // 缩小行高
                double lineLength = 40.0;
                int index = 0;

                foreach (var style in uniqueStyles.Values)
                {
                    Point3d start = new Point3d(insertPt.X, insertPt.Y - (index * rowHeight), 0);
                    Point3d end = new Point3d(insertPt.X + lineLength, start.Y, 0);

                    using (Line samplePath = new Line(start, end))
                    {
                        // ✨ 修复：传入 ed 参数，确保调用链完整
                        RenderPath(btr, tr, ed, samplePath, style, new ObjectIdCollection());

                        // 添加文字标注
                        MText mt = new MText();
                        mt.Contents = style.PathType == PathCategory.Solid ? "连续线样式" : $"阵列样式({style.SelectedBlockName})";
                        mt.Location = new Point3d(end.X + 5, end.Y, 0);
                        mt.Height = 2.5;
                        mt.Attachment = AttachmentPoint.MiddleLeft;
                        btr.AppendEntity(mt);
                        tr.AddNewlyCreatedDBObject(mt, true);
                    }
                    index++;
                }
                tr.Commit();
                ed.WriteMessage($"\n[成功] 已根据图中数据生成 {uniqueStyles.Count} 项图例。");
            }
        }

        private static Polyline ConvertToPolyline(Curve curve)
        {
            if (curve is Polyline pl) return pl.Clone() as Polyline;
            Polyline newPl = new Polyline();
            double len = curve.GetDistanceAtParameter(curve.EndParam);
            for (int i = 0; i <= 5; i++)
            {
                Point3d pt = curve.GetPointAtDist(len * (i / 5.0));
                newPl.AddVertexAt(i, new Point2d(pt.X, pt.Y), 0, 0, 0);
            }
            return newPl;
        }

        private static void RenderSkeleton(BlockTableRecord btr, Transaction tr, Curve curve)
        {
            Entity skeleton = (Entity)curve.Clone();
            skeleton.Color = Color.FromColorIndex(ColorMethod.ByAci, 253); // 淡灰色
            btr.AppendEntity(skeleton);
            tr.AddNewlyCreatedDBObject(skeleton, true);
        }

        /// <summary>
        /// 根据缩进参数修剪路径曲线。
        /// 修改逻辑：增加了对 headLen 和 tailLen 的范围限制（Clamping），
        /// 确保当缩进为负数时，距离不会小于 0 或大于总长，从而支持线体与箭头的重叠而不引发异常。
        /// </summary>
        private static Curve TrimPath(Curve rawPath, double headLen, double tailLen)
        {
            try
            {
                double totalLen = rawPath.GetDistanceAtParameter(rawPath.EndParam);

                // 计算实际的采样距离，确保在 [0, totalLen] 范围内
                // 正数缩进 = 向内缩短；负数或零 = 不缩短（由于 CAD 曲线扩展较复杂，负数暂按不缩短处理以保证重叠）
                double d1 = Math.Max(0, Math.Min(totalLen - 1e-4, tailLen));
                double d2 = Math.Max(d1 + 1e-4, Math.Min(totalLen, totalLen - headLen));

                if (totalLen <= (d1 + (totalLen - d2) + 1e-4)) return null;

                double p1 = rawPath.GetParameterAtDistance(d1);
                double p2 = rawPath.GetParameterAtDistance(d2);

                using (DBObjectCollection subCurves = rawPath.GetSplitCurves(new DoubleCollection { p1, p2 }))
                {
                    // 获取中间段曲线
                    int targetIndex = (subCurves.Count >= 3) ? 1 : 0;
                    return subCurves[targetIndex] as Curve;
                }
            }
            catch { return null; }
        }

        private static Curve CreatePathCurve(Point3dCollection pts, bool curved)
        {
            // 如果是曲线模式且点数充足，生成样条曲线
            if (curved && pts.Count > 2)
            {
                return new Spline(pts, 3, 0);
            }
            else
            {
                // ✨ 直线模式：确保 Polyline 正确添加了所有顶点
                Polyline pl = new Polyline { Plinegen = true };
                for (int i = 0; i < pts.Count; i++)
                {
                    pl.AddVertexAt(i, new Point2d(pts[i].X, pts[i].Y), 0, 0, 0);
                }
                return pl;
            }
        }
        private static double CalculateBezierWidth(double t, double s, double m, double e) => (1 - t) * (1 - t) * s + 2 * t * (1 - t) * m + t * t * e;

        private static void EnsureLayer(Database db, Transaction tr, AnalysisStyle s)
        {
            LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
            string name = string.IsNullOrEmpty(s.TargetLayer) ? "0" : s.TargetLayer;
            if (!lt.Has(name))
            {
                lt.UpgradeOpen();
                using (LayerTableRecord ltr = new LayerTableRecord { Name = name })
                {
                    lt.Add(ltr);
                    tr.AddNewlyCreatedDBObject(ltr, true);
                }
            }
        }
        /// <summary>
        /// 确保 CAD 数据库中已加载指定的线型（支持 acad.lin 和 acadiso.lin）。
        /// 改进点：增加了事务提交逻辑和多线型文件搜索，确保加载后立即生效。
        /// </summary>
        /// <summary>
        /// 确保 CAD 数据库中已加载指定的线型（如 ANALYSIS_DASH）。
        /// 改进点：
        /// 1. 增加了 doc.LockDocument() 逻辑，彻底解决异步调用下的 eLockViolation 错误。
        /// 2. 优化了路径获取逻辑，确保在各种安装环境下都能定位到 Assets 文件夹。
        /// </summary>
        private static ObjectId EnsureLinetypeLoaded(Document doc, string ltName)
        {
            Database db = doc.Database;
            Editor ed = doc.Editor;

            // 1. 先检查内存中是否已存在
            using (var tr = db.TransactionManager.StartTransaction())
            {
                LinetypeTable lt = (LinetypeTable)tr.GetObject(db.LinetypeTableId, OpenMode.ForRead);
                if (lt.Has(ltName)) return lt[ltName];
            }

            // 2. 定位插件目录
            string dllPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string isoLinPath = Path.Combine(dllPath, "Assets", "acadiso.lin");
            string stdLinPath = Path.Combine(dllPath, "Assets", "acad.lin");

            string finalLinPath = "";
            if (File.Exists(isoLinPath)) finalLinPath = isoLinPath;
            else if (File.Exists(stdLinPath)) finalLinPath = stdLinPath;

            // 3. 执行锁定并加载
            try
            {
                // ✨ 核心修复：异步环境下修改数据库必须锁定文档
                using (doc.LockDocument())
                {
                    if (!string.IsNullOrEmpty(finalLinPath))
                    {
                        db.LoadLineTypeFile(ltName, finalLinPath);
                    }
                    else
                    {
                        // 兜底方案：尝试从 CAD 系统路径加载
                        try { db.LoadLineTypeFile(ltName, "acadiso.lin"); }
                        catch { db.LoadLineTypeFile(ltName, "acad.lin"); }
                    }
                }

                // 4. 加载后重新获取 ID
                using (var trVerify = db.TransactionManager.StartTransaction())
                {
                    LinetypeTable ltVerify = (LinetypeTable)trVerify.GetObject(db.LinetypeTableId, OpenMode.ForRead);
                    if (ltVerify.Has(ltName)) return ltVerify[ltName];
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n[错误] 自动加载线型失败: {ex.Message}");
            }

            return db.ContinuousLinetype;
        }

    }
}