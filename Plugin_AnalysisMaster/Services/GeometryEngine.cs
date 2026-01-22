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
using Polyline = Autodesk.AutoCAD.DatabaseServices.Polyline;

namespace Plugin_AnalysisMaster.Services
{
    public static class GeometryEngine
    {

        private const string RegAppName = "ANALYSIS_MASTER_STYLE";
        // 存放每条路径预读取的数据，避免循环中操作数据库
        private class PathPlaybackData
        {
            public List<Point3d> Points { get; set; }
            public Autodesk.AutoCAD.Colors.Color Color { get; set; }
            public string Layer { get; set; }
            public ObjectId OriginalId { get; set; }
        }

        /// <summary>
        /// 异步播放序列逻辑
        /// </summary>
        public static async Task PlaySequenceAsync(
            List<AnimPathItem> items,
            double thickness,
            bool isLoop,
            bool isPersistent,
            CancellationToken token)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            Editor ed = doc.Editor;

            // 1. 预读取所有数据到内存
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

            // 记录所有已创建的瞬态对象，用于最后统一清理
            List<Polyline> activeTransients = new List<Polyline>();
            IntegerCollection vps = new IntegerCollection();

            try
            {
                do
                {
                    // A. 隐藏所有原始线条
                    ToggleEntitiesVisibility(items.Select(i => i.Id), false);

                    // 按组号分组并排序
                    var groupedData = items
                        .Select((item, index) => new { item, data = allData.FirstOrDefault(d => d.OriginalId == item.Id) })
                        .Where(x => x.data != null)
                        .GroupBy(x => x.item.GroupNumber)
                        .OrderBy(g => g.Key);

                    foreach (var group in groupedData)
                    {
                        if (token.IsCancellationRequested) return;

                        List<Polyline> currentGroupLines = new List<Polyline>();
                        var groupPaths = group.ToList();

                        // B. 初始化本组所有瞬态线（显示起点）
                        foreach (var path in groupPaths)
                        {
                            Polyline pl = new Polyline();
                            pl.Color = path.data.Color;
                            pl.Layer = path.data.Layer;
                            pl.ConstantWidth = thickness; // 应用演示加粗
                            pl.AddVertexAt(0, new Point2d(path.data.Points[0].X, path.data.Points[0].Y), 0, 0, 0);

                            TransientManager.CurrentTransientManager.AddTransient(pl, TransientDrawingMode.DirectTopmost, 128, vps);
                            currentGroupLines.Add(pl);
                            activeTransients.Add(pl);
                        }

                        // C. 组内同步生长循环
                        int maxSteps = groupPaths.Max(p => p.data.Points.Count);
                        for (int step = 1; step < maxSteps; step++)
                        {
                            if (token.IsCancellationRequested) return;

                            for (int i = 0; i < groupPaths.Count; i++)
                            {
                                var pathData = groupPaths[i].data;
                                if (step < pathData.Points.Count)
                                {
                                    currentGroupLines[i].AddVertexAt(step, new Point2d(pathData.Points[step].X, pathData.Points[step].Y), 0, 0, 0);
                                    TransientManager.CurrentTransientManager.UpdateTransient(currentGroupLines[i], vps);
                                }
                            }

                            Autodesk.AutoCAD.Internal.Utils.FlushGraphics();
                            await Task.Delay(20, token); // 使用带 Token 的延迟
                        }
                    }

                    if (!isLoop) break;

                    // 循环模式下，重新开始前清理一遍
                    ClearTransients(activeTransients, vps);
                    activeTransients.Clear();
                    await Task.Delay(500, token); // 循环间隔

                } while (isLoop && !token.IsCancellationRequested);
            }
            catch (OperationCanceledException)
            {
                // 捕获取消异常，不向外抛出
            }
            finally
            {
                // D. 最终清理
                ClearTransients(activeTransients, vps);

                // ✨ 核心修改：只有勾选了“保留迹线”，才把原始线恢复显示
                if (isPersistent)
                {
                    ToggleEntitiesVisibility(items.Select(i => i.Id), true);
                }
                else
                {
                    // 如果不保留，我们不仅不恢复显示，甚至可以提示用户是否需要删除它们
                    // 或者保持隐藏状态，让用户通过其他方式恢复
                }

                ed.UpdateScreen();
            }
        }

        // 辅助方法：读取实体数据
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
                // 解析指纹获取颜色
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

        // 辅助方法：批量切换可见性
        /// <summary>
        /// 批量切换实体可见性（增强版：支持自动隐藏关联组）
        /// </summary>
        private static void ToggleEntitiesVisibility(IEnumerable<ObjectId> ids, bool visible)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            using (doc.LockDocument())
            {
                using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
                {
                    // 使用 HashSet 防止重复处理同一个组员
                    HashSet<ObjectId> allToToggle = new HashSet<ObjectId>();

                    foreach (var id in ids)
                    {
                        if (id.IsNull || !id.IsValid) continue;
                        allToToggle.Add(id);

                        // ✨ 核心逻辑：检查该实体是否属于某个 Group (常见于阵列模式)
                        Entity ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                        if (ent != null)
                        {
                            ObjectIdCollection reactorIds = ent.GetPersistentReactorIds();
                            if (reactorIds != null)
                            {
                                foreach (ObjectId rId in reactorIds)
                                {
                                    // 检查反应器是否为 Group 类型
                                    if (rId.ObjectClass.IsDerivedFrom(RXClass.GetClass(typeof(Group))))
                                    {
                                        Group gp = tr.GetObject(rId, OpenMode.ForRead) as Group;
                                        if (gp != null)
                                        {
                                            // 将组内所有成员加入待操作列表
                                            foreach (ObjectId memberId in gp.GetAllEntityIds())
                                            {
                                                if (!allToToggle.Contains(memberId))
                                                    allToToggle.Add(memberId);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }

                    // 统一执行可见性切换
                    foreach (var toggleId in allToToggle)
                    {
                        Entity target = tr.GetObject(toggleId, OpenMode.ForWrite) as Entity;
                        if (target != null)
                        {
                            target.Visible = visible;
                        }
                    }

                    tr.Commit();
                }
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
                        // ✨ 修改：增加 ed 参数传递
                        RenderPath(btr, tr, ed, path, style, allIds);
                    }

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
                    ids.Add(AddToDb(btr, tr, pl, style, curve));
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
                        // ✨ 修改：传入 curve 以便 AddToDb 记录采样点
                        ids.Add(AddToDb(btr, tr, br, style, curve));
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
        /// 实体入库并附加样式指纹及动画采样点。
        /// 修改逻辑：增加了调试信息输出。
        /// 在保存动画数据时，会向命令行发送记录的采样点数量，方便确认数据是否已成功持久化。
        /// </summary>
        private static ObjectId AddToDb(BlockTableRecord btr, Transaction tr, Entity ent, AnalysisStyle style, Curve rawPath = null)
        {
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;

            // 1. 基础属性设置
            Database db = btr.Database;
            RegAppTable rat = (RegAppTable)tr.GetObject(db.RegAppTableId, OpenMode.ForRead);
            if (!rat.Has(RegAppName))
            {
                rat.UpgradeOpen();
                RegAppTableRecord ratr = new RegAppTableRecord { Name = RegAppName };
                rat.Add(ratr);
                tr.AddNewlyCreatedDBObject(ratr, true);
            }

            ent.Layer = style.TargetLayer;
            ent.Color = Autodesk.AutoCAD.Colors.Color.FromRgb(style.MainColor.R, style.MainColor.G, style.MainColor.B);

            // 2. 写入样式指纹 (XData)
            string styleFingerprint = string.Format("{0}|{1}|{2}|{3}|{4}|{5}|{6}",
                style.PathType, style.SelectedBlockName, style.SelectedBlockName2, style.IsComposite,
                style.MainColor.ToString(), style.StartArrowType, style.EndArrowType);

            ResultBuffer rbXData = new ResultBuffer(
                new TypedValue((int)DxfCode.ExtendedDataRegAppName, RegAppName),
                new TypedValue((int)DxfCode.ExtendedDataAsciiString, styleFingerprint)
            );
            ent.XData = rbXData;

            // 3. 先将实体添加到数据库
            ObjectId id = btr.AppendEntity(ent);
            tr.AddNewlyCreatedDBObject(ent, true);

            // 4. 实体入库后，存储动画路径数据
            if (style.IsAnimated && rawPath != null)
            {
                ent.CreateExtensionDictionary();
                DBDictionary dict = tr.GetObject(ent.ExtensionDictionary, OpenMode.ForWrite) as DBDictionary;

                using (Xrecord xRec = new Xrecord())
                {
                    using (ResultBuffer rbAnim = new ResultBuffer())
                    {
                        rbAnim.Add(new TypedValue((int)DxfCode.Text, styleFingerprint));
                        rbAnim.Add(new TypedValue((int)DxfCode.Real, style.SamplingInterval));

                        int ptCount = 0;
                        double totalLen = rawPath.GetDistanceAtParameter(rawPath.EndParam);
                        for (double d = 0; d <= totalLen; d += style.SamplingInterval)
                        {
                            rbAnim.Add(new TypedValue((int)DxfCode.XCoordinate, rawPath.GetPointAtDist(d)));
                            ptCount++;
                        }
                        rbAnim.Add(new TypedValue((int)DxfCode.XCoordinate, rawPath.EndPoint));
                        ptCount++;

                        xRec.Data = rbAnim;
                        dict.SetAt("ANALYSIS_ANIM_DATA", xRec);
                        tr.AddNewlyCreatedDBObject(xRec, true);

                        // ✨ 调试日志：确认保存成功
                        ed.WriteMessage($"\n[动画调试] 成功保存动画数据：采样点数 = {ptCount}, 采样间距 = {style.SamplingInterval}");
                    }
                }
            }
            else if (style.IsAnimated && rawPath == null)
            {
                ed.WriteMessage("\n[动画调试] 警告：开启动画记录但路径原始曲线(rawPath)为空。");
            }

            return id;
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