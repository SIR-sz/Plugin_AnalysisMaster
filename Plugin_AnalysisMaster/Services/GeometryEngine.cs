using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Plugin_AnalysisMaster.Models;
using System;
// 解决 2015 版本命名空间冲突
using Polyline = Autodesk.AutoCAD.DatabaseServices.Polyline;

namespace Plugin_AnalysisMaster.Services
{
    public static class GeometryEngine
    {
        /// <summary>
        /// 核心入口：参数化生成动线
        /// </summary>
        // ✨ 核心入口：实现“骨架+皮肤”的参数化生成
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

                    // 1. 生成样条控制曲线（骨架路径）
                    Curve path = CreatePathCurve(points, style.IsCurved);
                    using (path)
                    {
                        Vector3d headDir = path.GetFirstDerivative(path.EndParam).GetNormal();
                        Vector3d tailDir = path.GetFirstDerivative(path.StartParam).GetNormal().Negate();

                        // 2. 计算独立的头尾几何顶点
                        Point3dCollection headPts = CalculateFeaturePoints(path.EndPoint, headDir, style, true);
                        Point3dCollection tailPts = CalculateFeaturePoints(path.StartPoint, tailDir, style, false);

                        // 3. 计算物理缩进
                        double headIndent = (style.EndCapStyle == ArrowHeadType.None) ? 0 : style.ArrowSize * 0.8;
                        double tailIndent = (style.StartCapStyle == ArrowHeadType.None) ? 0 : style.ArrowSize * 0.4;

                        Curve trimmedPath = TrimPath(path, headIndent, tailIndent);

                        // 4. 执行“骨架与皮肤”渲染
                        if (trimmedPath != null)
                        {
                            RenderBody(btr, tr, trimmedPath, style);
                            if (trimmedPath != path) trimmedPath.Dispose();
                        }

                        // 5. 渲染独立的端头填充
                        if (headPts.Count > 0) RenderFeature(btr, tr, headPts, style);
                        if (tailPts.Count > 0) RenderFeature(btr, tr, tailPts, style);
                    }
                    tr.Commit();
                }
            }
        }

        #region 几何计算引擎 (Mathematics)

        private static Curve CreatePathCurve(Point3dCollection pts, bool isCurved)
        {
            if (isCurved && pts.Count > 2)
                return new Spline(pts, 3, 0);

            Polyline pl = new Polyline();
            for (int i = 0; i < pts.Count; i++)
                pl.AddVertexAt(i, new Point2d(pts[i].X, pts[i].Y), 0, 0, 0);
            return pl;
        }

        private static Point3dCollection CalculateFeaturePoints(Point3d basePt, Vector3d dir, AnalysisStyle style, bool isHead)
        {
            Point3dCollection pts = new Point3dCollection();
            Vector3d norm = dir.GetPerpendicularVector();
            double s = style.ArrowSize;

            if (isHead)
            {
                switch (style.HeadType)
                {
                    case ArrowHeadType.Basic:
                        pts.Add(basePt);
                        pts.Add(basePt - dir * s + norm * s * 0.4);
                        pts.Add(basePt - dir * s - norm * s * 0.4);
                        break;
                    case ArrowHeadType.SwallowTail:
                        double d = s * style.SwallowDepth;
                        pts.Add(basePt);
                        pts.Add(basePt - dir * s + norm * s * 0.5);
                        pts.Add(basePt - dir * d); // 内凹点
                        pts.Add(basePt - dir * s - norm * s * 0.5);
                        break;
                    case ArrowHeadType.Circle:
                        // 圆形由两个 1.0 凸度的段组成
                        Point3d center = basePt - dir * s * 0.5;
                        pts.Add(center + norm * s * 0.5);
                        pts.Add(center - norm * s * 0.5);
                        break;
                }
            }
            else // 尾部逻辑
            {
                switch (style.TailType)
                {
                    case ArrowTailType.Swallow:
                        double d = s * style.SwallowDepth;
                        pts.Add(basePt);
                        pts.Add(basePt - dir * s + norm * s * 0.5);
                        pts.Add(basePt - dir * d);
                        pts.Add(basePt - dir * s - norm * s * 0.5);
                        break;
                }
            }
            return pts;
        }

        private static Curve TrimPath(Curve rawPath, double headLen, double tailLen)
        {
            double totalLen = rawPath.GetDistanceAtParameter(rawPath.EndParam);
            if (totalLen <= headLen + tailLen) return null;

            double startDist = tailLen;
            double endDist = totalLen - headLen;

            double p1 = rawPath.GetParameterAtDistance(startDist);
            double p2 = rawPath.GetParameterAtDistance(endDist);

            DBObjectCollection subCurves = rawPath.GetSplitCurves(new DoubleCollection { p1, p2 });
            // 返回中间那段
            return subCurves.Count >= 2 ? subCurves[1] as Curve : subCurves[0] as Curve;
        }

        #endregion

        #region 实体渲染引擎 (Rendering)
        // ✨ 补全：渲染骨架虚线的方法
        private static void RenderSkeleton(BlockTableRecord btr, Transaction tr, Curve curve)
        {
            Entity skeleton = (Entity)curve.Clone();
            Database db = btr.Database;

            // 尝试设置线型，防止 eKeyNotFound
            LinetypeTable lt = (LinetypeTable)tr.GetObject(db.LinetypeTableId, OpenMode.ForRead);
            if (lt.Has("HIDDEN"))
            {
                skeleton.Linetype = "HIDDEN";
            }
            else if (lt.Has("DASHED"))
            {
                skeleton.Linetype = "DASHED";
            }

            // 设置骨架为淡灰色并半透明
            skeleton.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 253);
            skeleton.Transparency = new Autodesk.AutoCAD.Colors.Transparency(180);

            btr.AppendEntity(skeleton);
            tr.AddNewlyCreatedDBObject(skeleton, true);
        }
        // ✨ 替换 RenderBody：支持前后宽度不一的渐变线

        // ✨ 完善后的渲染方法：支持自适应线型加载与比例应用
        private static void RenderBody(BlockTableRecord btr, Transaction tr, Curve curve, AnalysisStyle style)
        {
            if (style.PathType == PathCategory.None) return;

            Database db = btr.Database;
            double totalLen = curve.GetDistanceAtParameter(curve.EndParam);
            double sampleStep = style.ArrowSize * 0.4;
            int numSegments = (int)Math.Max(totalLen / sampleStep, 40);

            using (Polyline visualSkin = new Polyline())
            {
                // 必须开启线型生成，否则高密度顶点会导致虚线失效
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

                // 虚线自适应处理
                if (style.PathType == PathCategory.Dashed)
                {
                    LinetypeTable lt = (LinetypeTable)tr.GetObject(db.LinetypeTableId, OpenMode.ForRead);
                    string ltName = "DASHED";

                    // 自动加载线型逻辑
                    if (!lt.Has(ltName))
                    {
                        try
                        {
                            // 尝试加载公制或英制线型库
                            try { db.LoadLineTypeFile(ltName, "acadiso.lin"); }
                            catch { db.LoadLineTypeFile(ltName, "acad.lin"); }
                        }
                        catch { /* 加载失败处理 */ }
                    }

                    if (lt.Has(ltName))
                    {
                        visualSkin.Linetype = ltName;
                        // ✨ 直接应用 UI 同步过来的比例
                        visualSkin.LinetypeScale = style.LinetypeScale;
                    }
                }

                AddToDb(btr, tr, visualSkin, style);
            }

            RenderSkeleton(btr, tr, curve);
        }
        // ✨ 新增私有辅助方法：计算二次方宽度插值
        private static double CalculateBezierWidth(double t, double start, double mid, double end)
        {
            double invT = 1.0 - t;
            return (invT * invT * start) + (2 * t * invT * mid) + (t * t * end);
        }
        private static void RenderFeature(BlockTableRecord btr, Transaction tr, Point3dCollection pts, AnalysisStyle style)
        {
            using (Polyline pl = new Polyline())
            {
                for (int i = 0; i < pts.Count; i++)
                {
                    // 处理圆形的特殊凸度
                    double bulge = (style.HeadType == ArrowHeadType.Circle) ? 1.0 : 0.0;
                    pl.AddVertexAt(i, new Point2d(pts[i].X, pts[i].Y), bulge, 0, 0);
                }
                pl.Closed = true;
                AddToDb(btr, tr, pl, style);
                CreateHatch(btr, tr, pl, style);
            }
        }

        private static void AddToDb(BlockTableRecord btr, Transaction tr, Entity ent, AnalysisStyle style)
        {
            ent.Layer = style.TargetLayer;
            ent.Color = Color.FromRgb(style.MainColor.R, style.MainColor.G, style.MainColor.B);
            ent.LineWeight = MapLineWeight(style.LineWeight);
            btr.AppendEntity(ent);
            tr.AddNewlyCreatedDBObject(ent, true);
        }

        private static void CreateHatch(BlockTableRecord btr, Transaction tr, Polyline boundary, AnalysisStyle style)
        {
            Hatch h = new Hatch();
            btr.AppendEntity(h);
            tr.AddNewlyCreatedDBObject(h, true);
            h.SetHatchPattern(HatchPatternType.PreDefined, "SOLID");
            h.Layer = boundary.Layer;
            h.Color = boundary.Color;
            h.Transparency = new Transparency((byte)style.Transparency);
            h.AppendLoop(HatchLoopTypes.Default, new ObjectIdCollection { boundary.ObjectId });
            h.EvaluateHatch(true);
        }

        private static LineWeight MapLineWeight(double w)
        {
            if (w < 0.1) return LineWeight.LineWeight005;
            if (w < 0.2) return LineWeight.LineWeight015;
            if (w < 0.4) return LineWeight.LineWeight030;
            return LineWeight.LineWeight050;
        }

        private static void EnsureLayer(Database db, Transaction tr, AnalysisStyle style)
        {
            LayerTable lt = tr.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable;
            if (!lt.Has(style.TargetLayer))
            {
                lt.UpgradeOpen();
                LayerTableRecord ltr = new LayerTableRecord { Name = style.TargetLayer };
                lt.Add(ltr);
                tr.AddNewlyCreatedDBObject(ltr, true);
            }
        }

        #endregion
    }
}