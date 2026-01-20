using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Plugin_AnalysisMaster.Models;
using System;
using System.Collections.Generic; // ✨ 新增
using System.IO;
using System.Linq;                // ✨ 新增
using System.Reflection;
using Polyline = Autodesk.AutoCAD.DatabaseServices.Polyline;

namespace Plugin_AnalysisMaster.Services
{
    public static class GeometryEngine
    {
        private const string RegAppName = "ANALYSIS_MASTER_STYLE";
        /// <summary>
        /// 动线绘制主入口：支持多点连续拾取与自动打组
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
                        // ✨ 修改：直接调用 RenderPath 封装逻辑
                        RenderPath(btr, tr, path, style, allIds);
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
        /// ✨ 渲染路径封装：统一处理修剪、主体渲染和端头渲染
        /// 解决 CS0103 报错，并供 DrawAnalysisLine 和 GenerateLegend 共同调用
        /// </summary>
        public static void RenderPath(BlockTableRecord btr, Transaction tr, Curve path, AnalysisStyle style, ObjectIdCollection allIds)
        {
            double indent = style.CapIndent;
            // 修剪路径（处理缩进）
            Curve trimmedPath = TrimPath(path, indent, indent);

            if (trimmedPath != null)
            {
                // 1. 渲染主体（实线或阵列）
                RenderBody(btr, tr, trimmedPath, style, allIds);

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
        /// 包含：
        /// 1. 实线 (Solid) 渲染：基于贝塞尔宽度算法生成带粗细变化的多段线。
        /// 2. 阵列 (Pattern) 渲染：基于两端留白居中算法，支持“组合模式”下的两个图块交替排布。
        /// </summary>
        private static void RenderBody(BlockTableRecord btr, Transaction tr, Curve curve, AnalysisStyle style, ObjectIdCollection ids)
        {
            double len = curve.GetDistanceAtParameter(curve.EndParam);

            // 分支 1：实线渲染
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
                    ids.Add(AddToDb(btr, tr, pl, style));
                }
            }
            // 分支 2：阵列渲染 (包含组合交替逻辑)
            else if (style.PathType == PathCategory.Pattern)
            {
                // 获取主图元 ID
                ObjectId blockId1 = GetOrImportBlock(style.SelectedBlockName, btr.Database, tr);
                // 获取组合图元 ID (未开启组合则使用主图元)
                ObjectId blockId2 = style.IsComposite ? GetOrImportBlock(style.SelectedBlockName2, btr.Database, tr) : blockId1;

                if (blockId1.IsNull) return;

                double spacing = style.PatternSpacing;
                if (spacing <= 0.001) spacing = 10.0;

                // 1. 计算能容纳的单元总数
                int count = (int)(len / spacing) + 1;
                // 2. 计算居中对齐的起始余量
                double occupiedLen = (count - 1) * spacing;
                double margin = (len - occupiedLen) / 2.0;

                for (int i = 0; i < count; i++)
                {
                    double currentDist = margin + (i * spacing);
                    if (currentDist < 0) currentDist = 0;
                    if (currentDist > len) currentDist = len;

                    // ✨ 核心逻辑：如果是组合模式，按索引奇偶性交替使用 blockId1 和 blockId2
                    ObjectId currentBlockId = (i % 2 != 0 && style.IsComposite) ? blockId2 : blockId1;
                    // 安全检查，防止副图元导入失败
                    if (currentBlockId.IsNull) currentBlockId = blockId1;

                    Point3d pt = curve.GetPointAtDist(currentDist);
                    double param = curve.GetParameterAtDistance(currentDist);
                    Vector3d tan = curve.GetFirstDerivative(param).GetNormal();

                    using (BlockReference br = new BlockReference(pt, currentBlockId))
                    {
                        br.Rotation = Math.Atan2(tan.Y, tan.X);
                        br.ScaleFactors = new Scale3d(style.PatternScale);
                        br.Color = Autodesk.AutoCAD.Colors.Color.FromRgb(
                            style.MainColor.R, style.MainColor.G, style.MainColor.B);

                        ids.Add(AddToDb(btr, tr, br, style));
                    }
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
        /// 修改后的 AddToDb：将路径单元、颜色、以及起终点箭头全部纳入指纹识别
        /// </summary>
        private static ObjectId AddToDb(BlockTableRecord btr, Transaction tr, Entity ent, AnalysisStyle style)
        {
            // 确保注册了 AppName
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
            // 确保设置了颜色
            ent.Color = Color.FromRgb(style.MainColor.R, style.MainColor.G, style.MainColor.B);

            // ✨ 重新梳理样式指纹逻辑：增加起终点箭头判断
            // 格式：路径模式|主块名|副块名|组合模式|颜色|起点箭头|终点箭头
            string styleFingerprint = string.Format("{0}|{1}|{2}|{3}|{4}|{5}|{6}",
                style.PathType,
                style.SelectedBlockName,
                style.SelectedBlockName2,
                style.IsComposite,
                style.MainColor.ToString(),
                style.StartArrowType,
                style.EndArrowType);

            ResultBuffer rb = new ResultBuffer(
                new TypedValue((int)DxfCode.ExtendedDataRegAppName, RegAppName),
                new TypedValue((int)DxfCode.ExtendedDataAsciiString, styleFingerprint)
            );
            ent.XData = rb;

            ObjectId id = btr.AppendEntity(ent);
            tr.AddNewlyCreatedDBObject(ent, true);
            return id;
        }
        /// <summary>
        /// 修改后的图例生成逻辑：考虑了箭头的种类和个数
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
                        // 解析新版指纹（包含 7 个部分）
                        string[] parts = fingerprint.Split('|');
                        if (parts.Length < 7) continue;

                        var style = new AnalysisStyle();
                        style.PathType = (PathCategory)System.Enum.Parse(typeof(PathCategory), parts[0]);
                        style.SelectedBlockName = parts[1];
                        style.SelectedBlockName2 = parts[2];
                        style.IsComposite = bool.Parse(parts[3]);

                        // 还原颜色
                        var colorStr = parts[4];
                        style.MainColor = (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(colorStr);

                        // ✨ 还原箭头设置
                        style.StartArrowType = parts[5];
                        style.EndArrowType = parts[6];

                        // 默认一些图例显示的比例
                        style.ArrowSize = 1.0;
                        style.PatternScale = 1.0;

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
                double rowHeight = 30.0; // 调高了行高，防止大箭头重叠
                double lineLength = 60.0;
                int index = 0;

                foreach (var style in uniqueStyles.Values)
                {
                    Point3d start = new Point3d(insertPt.X, insertPt.Y - (index * rowHeight), 0);
                    Point3d end = new Point3d(insertPt.X + lineLength, start.Y, 0);

                    using (Line samplePath = new Line(start, end))
                    {
                        // ✨ 自动根据还原的 style 绘制主体和对应的箭头
                        RenderPath(btr, tr, samplePath, style, new ObjectIdCollection());
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