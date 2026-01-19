using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Plugin_AnalysisMaster.Models;
using System;
using Polyline = Autodesk.AutoCAD.DatabaseServices.Polyline;

namespace Plugin_AnalysisMaster.Services
{
    public static class GeometryEngine
    {
        public static void DrawAnalysisLine(Point3dCollection points, AnalysisStyle style)
        {
            if (points == null || points.Count < 2) return;
            Document doc = Application.DocumentManager.MdiActiveDocument;
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
                        Vector3d headDir = path.GetFirstDerivative(path.EndParam).GetNormal();
                        Vector3d tailDir = path.GetFirstDerivative(path.StartParam).GetNormal().Negate();
                        double headIndent = (style.EndCapStyle == ArrowHeadType.None) ? 0 : style.ArrowSize * 0.8;
                        double tailIndent = (style.StartCapStyle == ArrowHeadType.None) ? 0 : style.ArrowSize * 0.4;

                        Curve trimmedPath = TrimPath(path, headIndent, tailIndent);
                        if (trimmedPath != null)
                        {
                            RenderBody(btr, tr, trimmedPath, style, allIds);
                            if (trimmedPath != path) trimmedPath.Dispose();
                        }

                        Point3dCollection headPts = CalculateFeaturePoints(path.EndPoint, headDir, style, true);
                        Point3dCollection tailPts = CalculateFeaturePoints(path.StartPoint, tailDir, style, false);
                        if (headPts.Count > 0) RenderFeature(btr, tr, headPts, style, allIds);
                        if (tailPts.Count > 0) RenderFeature(btr, tr, tailPts, style, allIds);
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

        // ✨ 核心补全：提供给 UI 使用的预览图元生成方法
        public static DBObjectCollection GeneratePreviewEntities(Point3d startPt, Point3d endPt, AnalysisStyle style)
        {
            DBObjectCollection entities = new DBObjectCollection();
            Point3dCollection pts = new Point3dCollection { startPt, endPt };
            Curve path = CreatePathCurve(pts, style.IsCurved);
            if (path == null) return entities;

            using (path)
            {
                Vector3d dir = (endPt - startPt).GetNormal();
                double headIndent = (style.EndCapStyle == ArrowHeadType.None) ? 0 : style.ArrowSize * 0.8;
                double tailIndent = (style.StartCapStyle == ArrowHeadType.None) ? 0 : style.ArrowSize * 0.4;
                Curve trimmedPath = TrimPath(path, headIndent, tailIndent);
                Curve renderPath = trimmedPath ?? path;

                Database db = HostApplicationServices.WorkingDatabase;
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    if (style.PathType == PathCategory.Solid)
                    {
                        Polyline pl = new Polyline { Plinegen = true };
                        pl.Color = Color.FromRgb(style.MainColor.R, style.MainColor.G, style.MainColor.B);
                        double len = renderPath.GetDistanceAtParameter(renderPath.EndParam);
                        for (int i = 0; i <= 20; i++)
                        {
                            double t = i / 20.0;
                            Point3d pt = renderPath.GetPointAtDist(len * t);
                            pl.AddVertexAt(i, new Point2d(pt.X, pt.Y), 0, 0, 0);
                            if (i < 20)
                            {
                                pl.SetStartWidthAt(i, CalculateBezierWidth(t, style.StartWidth, style.MidWidth, style.EndWidth));
                                pl.SetEndWidthAt(i, CalculateBezierWidth((i + 1) / 20.0, style.StartWidth, style.MidWidth, style.EndWidth));
                            }
                        }
                        entities.Add(pl);
                    }
                    else if (style.PathType == PathCategory.Dashed || style.PathType == PathCategory.Pattern)
                    {
                        double spacing = style.ArrowSize * 1.5 * style.LinetypeScale;
                        double len = renderPath.GetDistanceAtParameter(renderPath.EndParam);
                        ObjectId bid = (style.PathType == PathCategory.Dashed) ? GetOrCreateBuiltInBlock(style.SelectedBuiltInPattern, db, tr) : GetBlockIdByName(style.CustomBlockName, db, tr);
                        if (!bid.IsNull)
                        {
                            for (double d = 0; d <= len; d += spacing)
                            {
                                BlockReference br = new BlockReference(renderPath.GetPointAtDist(d), bid);
                                double t = d / len;
                                br.ScaleFactors = new Scale3d(CalculateBezierWidth(t, style.StartWidth, style.MidWidth, style.EndWidth));
                                br.Color = Color.FromRgb(style.MainColor.R, style.MainColor.G, style.MainColor.B);
                                entities.Add(br);
                            }
                        }
                    }
                    tr.Commit();
                }
                if (trimmedPath != null && trimmedPath != path) trimmedPath.Dispose();
            }
            return entities;
        }

        // ✨ 完整方法：实现“动态多段线单元”阵列渲染逻辑
        private static void RenderBody(BlockTableRecord btr, Transaction tr, Curve curve, AnalysisStyle style, ObjectIdCollection allIds)
        {
            if (style.PathType == PathCategory.None) return;

            double totalLen = curve.GetDistanceAtParameter(curve.EndParam);

            // 1. 实线类逻辑 (保持原样)
            if (style.PathType == PathCategory.Solid)
            {
                double sampleStep = style.ArrowSize * 0.4;
                int numSegments = (int)Math.Max(totalLen / sampleStep, 40);

                using (Polyline visualSkin = new Polyline())
                {
                    visualSkin.Plinegen = true;
                    for (int i = 0; i <= numSegments; i++)
                    {
                        double currentDist = (totalLen / (double)numSegments) * i;
                        Point3d pt = curve.GetPointAtDist(currentDist);
                        visualSkin.AddVertexAt(i, new Point2d(pt.X, pt.Y), 0, 0, 0);

                        if (i < numSegments)
                        {
                            double t = (double)i / (double)numSegments;
                            double nextT = (double)(i + 1) / (double)numSegments;
                            double currentW = CalculateBezierWidth(t, style.StartWidth, style.MidWidth, style.EndWidth);
                            double nextW = CalculateBezierWidth(nextT, style.StartWidth, style.MidWidth, style.EndWidth);
                            visualSkin.SetStartWidthAt(i, currentW);
                            visualSkin.SetEndWidthAt(i, nextW);
                        }
                    }
                    allIds.Add(AddToDb(btr, tr, visualSkin, style));
                }
            }
            // 2. ✨ 新增虚线类逻辑：动态生成多段线单元并阵列放置
            else if (style.PathType == PathCategory.Dashed)
            {
                // 计算单元长度和间距（受线形比例滑块控制）
                double dashUnitLen = style.ArrowSize * 0.8 * style.LinetypeScale;
                double gapLen = style.ArrowSize * 0.6 * style.LinetypeScale;
                double currentDist = 0;

                while (currentDist < totalLen)
                {
                    double segmentEnd = Math.Min(currentDist + dashUnitLen, totalLen);
                    if (segmentEnd <= currentDist) break;

                    // 获取当前段在路径上的参数区间
                    double pStart = curve.GetParameterAtDistance(currentDist);
                    double pEnd = curve.GetParameterAtDistance(segmentEnd);

                    // 提取骨架路径的子段
                    using (DBObjectCollection subCurves = curve.GetSplitCurves(new DoubleCollection { pStart, pEnd }))
                    {
                        Curve dashSeg = null;
                        // 获取切割后的目标段（逻辑：pStart后的第一段或中间段）
                        if (subCurves.Count >= 3) dashSeg = subCurves[1] as Curve;
                        else if (subCurves.Count >= 2) dashSeg = (currentDist == 0) ? subCurves[0] as Curve : subCurves[1] as Curve;
                        else dashSeg = subCurves[0] as Curve;

                        if (dashSeg != null)
                        {
                            // 将子段转换为多段线实体，以便应用用户定义的物理宽度
                            using (Polyline unitPl = ConvertToPolyline(dashSeg))
                            {
                                // ✨ 应用用户通过数值调整的起点和终点物理宽度
                                unitPl.SetStartWidthAt(0, style.StartWidth);
                                unitPl.SetEndWidthAt(unitPl.NumberOfVertices - 1, style.EndWidth);

                                // 添加到数据库并收集 ID
                                allIds.Add(AddToDb(btr, tr, unitPl, style));
                            }
                        }
                    }
                    // 移动到下一个阵列点
                    currentDist += dashUnitLen + gapLen;
                }
            }
            // 3. 自定义类逻辑 (保持原有的块阵列逻辑)
            else if (style.PathType == PathCategory.CustomPattern)
            {
                // (执行您之前已完成的 CustomPattern 块参照阵列代码)
            }

            RenderSkeleton(btr, tr, curve);
        }
        // 在 Services\GeometryEngine.cs 中添加
        private static void RenderSkeleton(BlockTableRecord btr, Transaction tr, Curve curve)
        {
            // 克隆原始路径作为骨架
            Entity skeleton = (Entity)curve.Clone();

            // 设置为淡灰色 (ACI 253) 并应用半透明效果
            skeleton.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 253);
            skeleton.Transparency = new Autodesk.AutoCAD.Colors.Transparency(180);

            // 如果图纸中有 HIDDEN 线型则应用，否则保持实线
            Database db = btr.Database;
            LinetypeTable lt = (LinetypeTable)tr.GetObject(db.LinetypeTableId, OpenMode.ForRead);
            if (lt.Has("HIDDEN"))
            {
                skeleton.Linetype = "HIDDEN";
            }

            btr.AppendEntity(skeleton);
            tr.AddNewlyCreatedDBObject(skeleton, true);
        }

        // ✨ 完整方法：辅助工具，将任意曲线子段转换为带权重的多段线单元
        private static Polyline ConvertToPolyline(Curve curve)
        {
            // 如果本身就是多段线，直接克隆
            if (curve is Polyline pl) return pl.Clone() as Polyline;

            Polyline newPl = new Polyline();
            double len = curve.GetDistanceAtParameter(curve.EndParam);

            // 为了保证虚线单元在弯曲路径上平滑，采样 6 个点即可
            int samples = 5;
            for (int i = 0; i <= samples; i++)
            {
                Point3d pt = curve.GetPointAtDist(len * (i / (double)samples));
                newPl.AddVertexAt(i, new Point2d(pt.X, pt.Y), 0, 0, 0);
            }
            return newPl;
        }

        private static ObjectId AddToDb(BlockTableRecord btr, Transaction tr, Entity ent, AnalysisStyle style)
        {
            ent.Layer = style.TargetLayer;
            ent.Color = Color.FromRgb(style.MainColor.R, style.MainColor.G, style.MainColor.B);
            ObjectId id = btr.AppendEntity(ent);
            tr.AddNewlyCreatedDBObject(ent, true);
            return id;
        }

        private static void RenderFeature(BlockTableRecord btr, Transaction tr, Point3dCollection pts, AnalysisStyle style, ObjectIdCollection idCol)
        {
            using (Polyline pl = new Polyline())
            {
                for (int i = 0; i < pts.Count; i++) pl.AddVertexAt(i, new Point2d(pts[i].X, pts[i].Y), 0, 0, 0);
                pl.Closed = true;
                idCol.Add(AddToDb(btr, tr, pl, style));
            }
        }

        private static ObjectId GetOrCreateBuiltInBlock(BuiltInPatternType type, Database db, Transaction tr)
        {
            string blockName = "_AM_INTERNAL_" + type.ToString();
            BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
            if (bt.Has(blockName)) return bt[blockName];
            bt.UpgradeOpen();
            using (BlockTableRecord btr = new BlockTableRecord { Name = blockName })
            {
                Entity shape = (type == BuiltInPatternType.Dot) ? (Entity)new Circle(Point3d.Origin, Vector3d.ZAxis, 0.5) : new Line(new Point3d(-0.5, 0, 0), new Point3d(0.5, 0, 0));
                btr.AppendEntity(shape); bt.Add(btr); tr.AddNewlyCreatedDBObject(btr, true);
                return btr.ObjectId;
            }
        }

        private static ObjectId GetBlockIdByName(string name, Database db, Transaction tr) => (string.IsNullOrEmpty(name)) ? ObjectId.Null : (((BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead)).Has(name) ? ((BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead))[name] : ObjectId.Null);
        private static Curve CreatePathCurve(Point3dCollection pts, bool isCurved) => (isCurved && pts.Count > 2) ? (Curve)new Spline(pts, 3, 0) : new Polyline();
        private static Point3dCollection CalculateFeaturePoints(Point3d basePt, Vector3d dir, AnalysisStyle style, bool isHead) => new Point3dCollection { basePt, basePt - dir * style.ArrowSize };
        private static Curve TrimPath(Curve raw, double headLen, double tailLen)
        {
            try
            {
                double totalLen = raw.GetDistanceAtParameter(raw.EndParam);

                // 增加冗余量检查，如果缩进总和大于等于总长度，直接返回 null
                if (totalLen <= (headLen + tailLen + 1e-3)) return null;

                double p1 = raw.GetParameterAtDistance(tailLen);
                double p2 = raw.GetParameterAtDistance(totalLen - headLen);

                using (DBObjectCollection subCurves = raw.GetSplitCurves(new DoubleCollection { p1, p2 }))
                {
                    // 确保集合不为空且元素数量足够
                    if (subCurves != null && subCurves.Count > 0)
                    {
                        // 正常的两刀切割会产生 3 段，索引 1 为中间路径
                        // 如果只产生 2 段或 1 段，则安全取值防止越界
                        int targetIndex = (subCurves.Count >= 3) ? 1 : 0;
                        return subCurves[targetIndex].Clone() as Curve;
                    }
                }
            }
            catch
            {
                // 捕捉可能的几何计算异常，防止 CAD 直接崩溃
            }
            return null;
        }
        private static double CalculateBezierWidth(double t, double s, double m, double e) => (1 - t) * (1 - t) * s + 2 * t * (1 - t) * m + t * t * e;
        // ✨ 修复：确保图层存在，防止 eKeyNotFound 崩溃
        private static void EnsureLayer(Database db, Transaction tr, AnalysisStyle style)
        {
            // 安全检查：如果图层名为空，回退到 0 图层
            if (string.IsNullOrEmpty(style.TargetLayer))
                style.TargetLayer = "0";

            LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);

            if (!lt.Has(style.TargetLayer))
            {
                lt.UpgradeOpen();
                using (LayerTableRecord ltr = new LayerTableRecord())
                {
                    ltr.Name = style.TargetLayer;
                    lt.Add(ltr);
                    tr.AddNewlyCreatedDBObject(ltr, true);
                }
            }
        }
    }
}