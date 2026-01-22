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

        private const string RegAppName = "ANALYSIS_MASTER_STYLE";
        // 修改位置：GeometryEngine.cs 约第 26 行
        private class PathPlaybackData
        {
            public List<Point3d> Points { get; set; }
            public Autodesk.AutoCAD.Colors.Color Color { get; set; }
            public string Layer { get; set; }
            public ObjectId OriginalId { get; set; }
            // ✨ 简化：删除了 SortedMembers 和 LastShownMemberIndex
        }

        /// <summary>
        /// 异步播放序列逻辑
        /// </summary>
        // 位置：GeometryEngine.cs 约第 40 行
        // 修改位置：GeometryEngine.cs -> PlaySequenceAsync
        public static async Task PlaySequenceAsync(
              List<AnimPathItem> items,
              double thickness,
              int speedMultiplier,
              bool isLoop,
              bool isPersistent,
              CancellationToken token)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            Editor ed = doc.Editor;

            List<PathPlaybackData> allData = new List<PathPlaybackData>();
            using (doc.LockDocument())
            using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
            {
                foreach (var item in items)
                {
                    var data = FetchPathData(tr, item.Id);
                    if (data != null) allData.Add(data);
                }
                tr.Commit();
            }

            if (allData.Count == 0) return;
            List<Polyline> activeTransients = new List<Polyline>();
            IntegerCollection vps = new IntegerCollection();

            try
            {
                do
                {
                    // A. 隐藏所有原始实体（包括阵列单元）
                    ToggleEntitiesVisibility(items.Select(i => i.Id), false);

                    var groupedData = items
                        .Select(x => new { item = x, data = allData.FirstOrDefault(d => d.OriginalId == x.Id) })
                        .Where(x => x.data != null)
                        .GroupBy(x => x.item.GroupNumber)
                        .OrderBy(g => g.Key);

                    foreach (var group in groupedData)
                    {
                        if (token.IsCancellationRequested) return;

                        List<Polyline> currentGroupLines = new List<Polyline>();
                        var groupPaths = group.ToList();
                        int[] lastAddedIndices = new int[groupPaths.Count];

                        // B. 初始化引导线
                        foreach (var path in groupPaths)
                        {
                            Polyline pl = new Polyline { Plinegen = true };
                            pl.Color = path.data.Color;
                            pl.Layer = path.data.Layer;
                            pl.AddVertexAt(0, new Point2d(path.data.Points[0].X, path.data.Points[0].Y), 0, 0, 0);
                            pl.ConstantWidth = thickness;
                            TransientManager.CurrentTransientManager.AddTransient(pl, TransientDrawingMode.DirectTopmost, 128, vps);
                            currentGroupLines.Add(pl);
                            activeTransients.Add(pl);
                        }

                        // C. 引导线生长循环
                        for (int currentStep = speedMultiplier; ; currentStep += speedMultiplier)
                        {
                            if (token.IsCancellationRequested) return;
                            bool allFinished = true;

                            for (int i = 0; i < groupPaths.Count; i++)
                            {
                                var data = groupPaths[i].data;
                                var pl = currentGroupLines[i];
                                if (lastAddedIndices[i] >= data.Points.Count - 1) continue;

                                int nextPtIdx = Math.Min(currentStep, data.Points.Count - 1);
                                if (nextPtIdx > lastAddedIndices[i])
                                {
                                    int vtxIdx = pl.NumberOfVertices;
                                    pl.AddVertexAt(vtxIdx, new Point2d(data.Points[nextPtIdx].X, data.Points[nextPtIdx].Y), 0, 0, 0);
                                    lastAddedIndices[i] = nextPtIdx;
                                    TransientManager.CurrentTransientManager.UpdateTransient(pl, vps);
                                }
                                if (lastAddedIndices[i] < data.Points.Count - 1) allFinished = false;
                            }

                            Autodesk.AutoCAD.Internal.Utils.FlushGraphics();
                            await Task.Delay(10, token);
                            if (allFinished) break;
                        }

                        // ✨ D. 简化核心：本组引导线长完后，立即显示该组对应的所有原始单元
                        // 这里的 ToggleEntitiesVisibility 会通过组关联自动显示所有阵列块
                        ToggleEntitiesVisibility(groupPaths.Select(p => p.item.Id), true);

                        // 此时可以移除当前组的引导线，防止重叠
                        foreach (var pl in currentGroupLines)
                        {
                            TransientManager.CurrentTransientManager.EraseTransient(pl, vps);
                            activeTransients.Remove(pl);
                            pl.Dispose();
                        }
                    }
                    if (!isLoop) break;
                    await Task.Delay(500, token);
                } while (isLoop && !token.IsCancellationRequested);
            }
            finally
            {
                ClearTransients(activeTransients, vps);
                ToggleEntitiesVisibility(items.Select(i => i.Id), true);
                ed.UpdateScreen();
            }
        }

        // 辅助方法：读取实体数据
        // 位置：GeometryEngine.cs 约第 105 行
        // 修改位置：GeometryEngine.cs 约第 105 行
        private static PathPlaybackData FetchPathData(Transaction tr, ObjectId id)
        {
            Entity ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
            if (ent == null || ent.ExtensionDictionary.IsNull) return null;

            DBDictionary dict = tr.GetObject(ent.ExtensionDictionary, OpenMode.ForRead) as DBDictionary;
            if (!dict.Contains("ANALYSIS_ANIM_DATA")) return null;

            Xrecord xRec = tr.GetObject(dict.GetAt("ANALYSIS_ANIM_DATA"), OpenMode.ForRead) as Xrecord;
            using (ResultBuffer rb = xRec.Data)
            {
                TypedValue[] dataArr = rb.AsArray();
                string fingerprint = dataArr[0].Value.ToString();
                string[] parts = fingerprint.Split('|');
                var colorStr = parts[4];
                var mediaColor = (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(colorStr);

                List<Point3d> pts = new List<Point3d>();
                for (int i = 2; i < dataArr.Length; i++)
                {
                    if (dataArr[i].TypeCode == (int)DxfCode.XCoordinate)
                        pts.Add((Point3d)dataArr[i].Value);
                }

                return new PathPlaybackData
                {
                    Points = pts,
                    Color = Autodesk.AutoCAD.Colors.Color.FromRgb(mediaColor.R, mediaColor.G, mediaColor.B),
                    Layer = ent.Layer,
                    OriginalId = id
                };
            }
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
        /// 批量切换实体可见性（增强版：增加安全性校验，防止 eInvalidInput）
        /// </summary>
       // 辅助：支持 Group 自动隐藏的可见性控制
        private static void ToggleEntitiesVisibility(IEnumerable<ObjectId> ids, bool visible)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
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
        }


        /// <summary>
        /// 动线绘制主入口：支持多点连续拾取与自动打组。
        /// 修改逻辑：将 Editor 对象显式传递给下级渲染方法以支持动画。
        /// </summary>
        public static void DrawAnalysisLine(Point3dCollection inputPoints, AnalysisStyle style)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Point3dCollection points = inputPoints;

            // 1. 交互式点位获取
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
                        jig.GetPoints().Add(jig.LastPoint);
                    }
                    else if (res.Status == PromptStatus.None)
                    {
                        jig.GetPoints().Add(jig.LastPoint);
                        break;
                    }
                    else
                    {
                        if (jig.GetPoints().Count < 2) return;
                        break;
                    }
                }
                points = jig.GetPoints();
            }

            if (points == null || points.Count < 2) return;

            // 2. 执行数据库写入
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
                        // 1. 渲染几何体（此时内部不采样，很快）
                        RenderPath(btr, tr, ed, path, style, allIds);

                        // 2. ✨ 核心修复：执行一次“骨架线采样”，并挂在第一个实体上
                        if (style.IsAnimated && allIds.Count > 0)
                        {
                            // 获取骨架线的采样点
                            List<Point3d> skeletonPoints = SampleCurve(path, style.SamplingInterval);

                            // 开启写模式打开第一个实体
                            Entity headEnt = tr.GetObject(allIds[0], OpenMode.ForWrite) as Entity;
                            if (headEnt != null)
                            {
                                // 写入指纹（包含 Guid 和采样点）
                                WriteFingerprint(tr, headEnt, style, skeletonPoints);
                            }
                        }
                    }

                    // 3. 打组（保持不变）
                    if (allIds.Count > 1)
                    {
                        DBDictionary gd = (DBDictionary)tr.GetObject(db.GroupDictionaryId, OpenMode.ForWrite);
                        Group grp = new Group("分析动线单元", true);
                        gd.SetAt("*", grp);
                        tr.AddNewlyCreatedDBObject(grp, true);
                        grp.Append(allIds);
                    }
                    tr.Commit();
                }
            }
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
                double step = Math.Max(style.ArrowSize * 0.15, 0.5);
                int samples = (int)Math.Max(len / step, 100);
                if (samples > 3000) samples = 3000;

                using (Polyline pl = new Polyline { Plinegen = true })
                {
                    for (int i = 0; i <= samples; i++)
                    {
                        double t = i / (double)samples;
                        Point3d pt = curve.GetPointAtDist(len * t);
                        pl.AddVertexAt(i, new Point2d(pt.X, pt.Y), 0, 0, 0);

                        if (i < samples)
                        {
                            double startW = CalculateBezierWidth(t, style.StartWidth, style.MidWidth, style.EndWidth);
                            double endW = CalculateBezierWidth((i + 1.0) / samples, style.StartWidth, style.MidWidth, style.EndWidth);
                            pl.SetStartWidthAt(i, startW);
                            pl.SetEndWidthAt(i, endW);
                        }
                    }
                    // ✨ 修改：传入 curve 以便 AddToDb 记录采样点
                    ids.Add(AddToDb(btr, tr, pl, style, null));
                }
            }
            else if (style.PathType == PathCategory.Pattern)
            {
                ObjectId blockId1 = GetOrImportBlock(style.SelectedBlockName, btr.Database, tr);
                ObjectId blockId2 = style.IsComposite ? GetOrImportBlock(style.SelectedBlockName2, btr.Database, tr) : blockId1;
                if (blockId1.IsNull) return;

                double spacing = style.PatternSpacing;
                if (spacing <= 0.001) spacing = 10.0;

                int count = (int)(len / spacing) + 1;
                double margin = (len - (count - 1) * spacing) / 2.0;

                for (int i = 0; i < count; i++)
                {
                    double currentDist = Math.Max(0, Math.Min(len, margin + (i * spacing)));
                    ObjectId currentBlockId = (i % 2 != 0 && style.IsComposite) ? blockId2 : blockId1;
                    if (currentBlockId.IsNull) currentBlockId = blockId1;

                    Point3d pt = curve.GetPointAtDist(currentDist);
                    Vector3d tan = curve.GetFirstDerivative(curve.GetParameterAtDistance(currentDist)).GetNormal();

                    using (BlockReference br = new BlockReference(pt, currentBlockId))
                    {
                        br.Rotation = Math.Atan2(tan.Y, tan.X);
                        br.ScaleFactors = new Scale3d(style.PatternScale);
                        // ✨ 核心修复：此处传 null。不再为每个图块重复采样。
                        ids.Add(AddToDb(btr, tr, br, style, null));
                    }
                }
            }
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
        /// 彻底解决阵列模式卡顿问题
        /// </summary>
        // --- 1. 给列表用的（新方法，处理阵列并打组） ---
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
                    if (ent.ObjectId.IsNull) // 只有没加入数据库的才添加
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

                // 仅在第一个实体写指纹
                WriteFingerprint(tr, entities[0], style, sampledPoints);
                tr.Commit();
            }
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
        private static void WriteFingerprint(Transaction tr, Entity ent, AnalysisStyle style, List<Point3d> sampledPoints)
        {
            // 如果没有采样点，则不写入动画数据
            if (!style.IsAnimated || sampledPoints == null || sampledPoints.Count == 0) return;

            if (ent.ExtensionDictionary.IsNull)
            {
                ent.UpgradeOpen();
                ent.CreateExtensionDictionary();
            }

            DBDictionary dict = (DBDictionary)tr.GetObject(ent.ExtensionDictionary, OpenMode.ForWrite);

            // 生成包含 GUID 的唯一指纹
            string fingerprint = $"{style.PathType}|{style.IsCurved}|{style.SelectedBlockName}|{DateTime.Now.Ticks}|{style.MainColor}|{Guid.NewGuid()}";

            using (Xrecord xRec = new Xrecord())
            {
                ResultBuffer rb = new ResultBuffer();
                rb.Add(new TypedValue((int)DxfCode.Text, fingerprint));
                rb.Add(new TypedValue((int)DxfCode.Int32, sampledPoints.Count));

                foreach (Point3d pt in sampledPoints)
                {
                    rb.Add(new TypedValue((int)DxfCode.XCoordinate, pt));
                }
                xRec.Data = rb;
                dict.SetAt("ANALYSIS_ANIM_DATA", xRec);
                tr.AddNewlyCreatedDBObject(xRec, true);
            }
        }
        /// <summary>
        /// 图例生成逻辑。
        /// 修改逻辑：由于 RenderPath 签名变更，需要在此处传递当前文档的 Editor 对象。
        /// </summary>
        public static void GenerateLegend(Document doc, string targetLayer)
        {
            Editor ed = doc.Editor;
            Database db = doc.Database;

            // 1. 框选图元
            SelectionFilter filter = new SelectionFilter(new TypedValue[] {
                new TypedValue(0, "LWPOLYLINE,POLYLINE,SPLINE,INSERT"),
                new TypedValue(8, targetLayer)
            });

            PromptSelectionResult selRes = ed.GetSelection(filter);
            if (selRes.Status != PromptStatus.OK) return;

            // 2. 识别并归类样式
            var uniqueStyles = new Dictionary<string, AnalysisStyle>();

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                foreach (SelectedObject selObj in selRes.Value)
                {
                    Entity ent = tr.GetObject(selObj.ObjectId, OpenMode.ForRead) as Entity;
                    if (ent == null || ent.XData == null) continue;

                    TypedValue[] xdata = ent.XData.AsArray();
                    var styleValue = xdata.FirstOrDefault(tv => tv.TypeCode == (int)DxfCode.ExtendedDataAsciiString);
                    if (styleValue.Value == null) continue;

                    string fingerprint = styleValue.Value.ToString();
                    if (!uniqueStyles.ContainsKey(fingerprint))
                    {
                        string[] parts = fingerprint.Split('|');
                        if (parts.Length < 7) continue;

                        var style = new AnalysisStyle();
                        style.PathType = (PathCategory)System.Enum.Parse(typeof(PathCategory), parts[0]);
                        style.SelectedBlockName = parts[1];
                        style.SelectedBlockName2 = parts[2];
                        style.IsComposite = bool.Parse(parts[3]);

                        var colorStr = parts[4];
                        style.MainColor = (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(colorStr);

                        style.StartArrowType = parts[5];
                        style.EndArrowType = parts[6];

                        style.ArrowSize = 1.0;
                        style.PatternScale = 1.0;
                        style.IsAnimated = false; // 图例生成不需要动画

                        uniqueStyles.Add(fingerprint, style);
                    }
                }

                if (uniqueStyles.Count == 0) return;

                // 3. 指定放置位置
                PromptPointResult ppr = ed.GetPoint("\n请指定图例放置起点: ");
                if (ppr.Status != PromptStatus.OK) return;

                Point3d insertPt = ppr.Value;
                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

                // 4. 绘制图例（一列排放）
                double rowHeight = 30.0;
                double lineLength = 60.0;
                int index = 0;

                foreach (var style in uniqueStyles.Values)
                {
                    Point3d start = new Point3d(insertPt.X, insertPt.Y - (index * rowHeight), 0);
                    Point3d end = new Point3d(insertPt.X + lineLength, start.Y, 0);

                    using (Line samplePath = new Line(start, end))
                    {
                        // ✨ 修改：增加 ed 参数传递
                        RenderPath(btr, tr, ed, samplePath, style, new ObjectIdCollection());
                    }
                    index++;
                }
                tr.Commit();
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

    }
}