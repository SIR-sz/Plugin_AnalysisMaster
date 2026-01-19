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

            // ✨ 修复 eDegenerateGeometry：通过预检路径总长度，拦截因点重合导致的退化几何
            double checkDist = 0;
            for (int i = 0; i < points.Count - 1; i++)
            {
                checkDist += points[i].DistanceTo(points[i + 1]);
            }

            // 如果总长度接近于 0，则不执行后续渲染，避免 EndParam 崩溃
            if (checkDist < 1e-6) return;

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
        // ✨ 完整方法：实现实线渐变、固定宽2的虚线阵列及自定义块阵列
        private static void RenderBody(BlockTableRecord btr, Transaction tr, Curve curve, AnalysisStyle style, ObjectIdCollection allIds)
        {
            if (style.PathType == PathCategory.None) return;

            Database db = btr.Database;
            double totalLen = curve.GetDistanceAtParameter(curve.EndParam);

            // 1. 实线类逻辑 (Solid) - 连续渐变宽度多段线
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
            // 2. 虚线类逻辑 (Dashed) - ✨ 改造点：宽度固定为 2 的多段线单元整列
            else if (style.PathType == PathCategory.Dashed)
            {
                // 线元长度与空隙：受 LinetypeScale 比例影响
                double dashLen = style.ArrowSize * 0.8 * style.LinetypeScale;
                double gapLen = style.ArrowSize * 0.6 * style.LinetypeScale;
                double currentDist = 0;

                while (currentDist < totalLen)
                {
                    double segmentEnd = Math.Min(currentDist + dashLen, totalLen);
                    if (segmentEnd <= currentDist) break;

                    double pStart = curve.GetParameterAtDistance(currentDist);
                    double pEnd = curve.GetParameterAtDistance(segmentEnd);

                    using (DBObjectCollection subCurves = curve.GetSplitCurves(new DoubleCollection { pStart, pEnd }))
                    {
                        // 获取切割后的路径子段
                        Curve dashSeg = (subCurves.Count >= 3) ? subCurves[1] as Curve : subCurves[0] as Curve;
                        if (dashSeg != null)
                        {
                            using (Polyline unitPl = ConvertToPolyline(dashSeg))
                            {
                                // ✨ 按照要求：宽度写死为 2.0
                                unitPl.SetStartWidthAt(0, 2.0);
                                unitPl.SetEndWidthAt(unitPl.NumberOfVertices - 1, 2.0);

                                allIds.Add(AddToDb(btr, tr, unitPl, style));
                            }
                        }
                        // 释放切割产生的中间对象
                        foreach (DBObject obj in subCurves) { if (obj != null && !obj.IsDisposed) obj.Dispose(); }
                    }
                    currentDist += dashLen + gapLen;
                }
            }
            // 3. 自定义类逻辑 (Pattern) - 种子块自适应阵列
            else if (style.PathType == PathCategory.Pattern)
            {
                double spacing = style.ArrowSize * 1.5 * style.LinetypeScale;
                double currentDist = 0;

                ObjectId patternBlockId = GetBlockIdByName(style.CustomBlockName, db, tr);

                if (!patternBlockId.IsNull)
                {
                    while (currentDist <= totalLen)
                    {
                        Point3d pt = curve.GetPointAtDist(currentDist);
                        double param = curve.GetParameterAtDistance(currentDist);
                        Vector3d tangent = curve.GetFirstDerivative(param).GetNormal();

                        // 块的大小依然受贝塞尔宽度趋势控制
                        double t = currentDist / totalLen;
                        double scale = CalculateBezierWidth(t, style.StartWidth, style.MidWidth, style.EndWidth);

                        using (BlockReference br = new BlockReference(pt, patternBlockId))
                        {
                            br.Rotation = Math.Atan2(tangent.Y, tangent.X);
                            br.ScaleFactors = new Scale3d(scale, scale, scale);
                            br.Color = Autodesk.AutoCAD.Colors.Color.FromRgb(style.MainColor.R, style.MainColor.G, style.MainColor.B);

                            ObjectId id = btr.AppendEntity(br);
                            tr.AddNewlyCreatedDBObject(br, true);
                            allIds.Add(id);
                        }
                        currentDist += spacing;
                    }
                }
            }

            RenderSkeleton(btr, tr, curve);
        }
        // 在 Services\GeometryEngine.cs 中添加
        // ✨ 完整方法：渲染底部灰色骨架线
        private static void RenderSkeleton(BlockTableRecord btr, Transaction tr, Curve curve)
        {
            Entity skeleton = (Entity)curve.Clone();
            // 使用淡灰色 (ACI 253)
            skeleton.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 253);
            btr.AppendEntity(skeleton);
            tr.AddNewlyCreatedDBObject(skeleton, true);
        }

        // ✨ 新增辅助方法：曲线转多段线
        // ✨ 完整方法：将路径子段转换为多段线，以便应用物理宽度
        private static Polyline ConvertToPolyline(Curve curve)
        {
            if (curve is Polyline pl) return pl.Clone() as Polyline;

            Polyline newPl = new Polyline();
            double len = curve.GetDistanceAtParameter(curve.EndParam);

            // 采样 5 个点以确保子段能够顺滑匹配原路径的弯曲度
            int samples = 5;
            for (int i = 0; i <= samples; i++)
            {
                Point3d pt = curve.GetPointAtDist(len * (i / (double)samples));
                newPl.AddVertexAt(i, new Point2d(pt.X, pt.Y), 0, 0, 0);
            }
            return newPl;
        }

        // ✨ 修复：将 AddToDb 返回类型改为 ObjectId
        private static ObjectId AddToDb(BlockTableRecord btr, Transaction tr, Entity ent, AnalysisStyle style)
        {
            ent.Layer = style.TargetLayer;
            ent.Color = Color.FromRgb(style.MainColor.R, style.MainColor.G, style.MainColor.B);
            ent.LineWeight = MapLineWeight(style.LineWeight);

            ObjectId id = btr.AppendEntity(ent);
            tr.AddNewlyCreatedDBObject(ent, true);
            return id;
        }
        // ✨ 修复：补全缺失的线宽映射方法
        private static Autodesk.AutoCAD.DatabaseServices.LineWeight MapLineWeight(double w)
        {
            if (w < 0.1) return Autodesk.AutoCAD.DatabaseServices.LineWeight.LineWeight005;
            if (w < 0.2) return Autodesk.AutoCAD.DatabaseServices.LineWeight.LineWeight015;
            if (w < 0.35) return Autodesk.AutoCAD.DatabaseServices.LineWeight.LineWeight030; // 0.3mm
            if (w < 0.5) return Autodesk.AutoCAD.DatabaseServices.LineWeight.LineWeight040;
            return Autodesk.AutoCAD.DatabaseServices.LineWeight.LineWeight050;
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

        // ✨ 完整方法：通过名称获取块 ID
        private static ObjectId GetBlockIdByName(string blockName, Database db, Transaction tr)
        {
            if (string.IsNullOrEmpty(blockName)) return ObjectId.Null;
            BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
            return bt.Has(blockName) ? bt[blockName] : ObjectId.Null;
        }
        private static Curve CreatePathCurve(Point3dCollection pts, bool isCurved) => (isCurved && pts.Count > 2) ? (Curve)new Spline(pts, 3, 0) : new Polyline();
        private static Point3dCollection CalculateFeaturePoints(Point3d basePt, Vector3d dir, AnalysisStyle style, bool isHead) => new Point3dCollection { basePt, basePt - dir * style.ArrowSize };
        private static Curve TrimPath(Curve rawPath, double headLen, double tailLen)
        {
            // ✨ 修复：增加长度预检与异常捕获，防止访问无效曲线参数
            double totalLen;
            try
            {
                totalLen = rawPath.GetDistanceAtParameter(rawPath.EndParam);
            }
            catch
            {
                // 如果曲线无法获取参数，说明几何已退化，直接返回 null
                return null;
            }

            // 增加冗余量检查，如果缩进总和大于总长度，返回 null
            if (totalLen <= (headLen + tailLen + 1e-4)) return null;

            double startDist = tailLen;
            double endDist = totalLen - headLen;

            double p1 = rawPath.GetParameterAtDistance(startDist);
            double p2 = rawPath.GetParameterAtDistance(endDist);

            using (DBObjectCollection subCurves = rawPath.GetSplitCurves(new DoubleCollection { p1, p2 }))
            {
                // 索引访问保护
                if (subCurves != null && subCurves.Count > 0)
                {
                    int targetIndex = (subCurves.Count >= 3) ? 1 : 0;
                    return subCurves[targetIndex] as Curve;
                }
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