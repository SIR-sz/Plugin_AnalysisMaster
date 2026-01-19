using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Plugin_AnalysisMaster.Models;
using System;
using System.IO;
using System.Reflection;
using Polyline = Autodesk.AutoCAD.DatabaseServices.Polyline;

namespace Plugin_AnalysisMaster.Services
{
    public static class GeometryEngine
    {
        /// <summary>
        /// 动线绘制主入口：支持多点连续拾取与自动打组
        /// </summary>
        public static void DrawAnalysisLine(Point3dCollection inputPoints, AnalysisStyle style)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Point3dCollection points = inputPoints;

            // 1. 交互式点位获取 (实现连续点击逻辑)
            if (points == null)
            {
                PromptPointOptions ppo = new PromptPointOptions("\n指定起点: ");
                PromptPointResult ppr = ed.GetPoint(ppo);
                if (ppr.Status != PromptStatus.OK) return;

                // 初始化预览 Jig
                var jig = new AnalysisLineJig(ppr.Value, style);

                // ✨ 核心逻辑：循环拾取点位，直到按下 Enter
                while (true)
                {
                    var res = ed.Drag(jig);
                    if (res.Status == PromptStatus.OK)
                    {
                        // 用户点击左键：确定当前点，并继续下一段预览
                        jig.GetPoints().Add(jig.LastPoint);
                    }
                    else if (res.Status == PromptStatus.None)
                    {
                        // 用户按回车或空格：结束拾取，加入最后一个预览点
                        jig.GetPoints().Add(jig.LastPoint);
                        break;
                    }
                    else
                    {
                        // 取消操作
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
                        // 计算端头缩进，防止线体穿透箭头
                        double headIndent = (style.EndCapStyle == ArrowHeadType.None) ? 0 : style.ArrowSize * 0.8;
                        double tailIndent = (style.StartCapStyle == ArrowHeadType.None) ? 0 : style.ArrowSize * 0.4;

                        Curve trimmedPath = TrimPath(path, headIndent, tailIndent);
                        if (trimmedPath != null)
                        {
                            RenderBody(btr, tr, trimmedPath, style, allIds);
                            if (trimmedPath != path) trimmedPath.Dispose();
                        }

                        // 渲染端头特征 (箭头或圆点)
                        if (style.EndCapStyle != ArrowHeadType.None)
                        {
                            Vector3d headDir = path.GetFirstDerivative(path.EndParam).GetNormal();
                            Point3dCollection headPts = CalculateFeaturePoints(path.EndPoint, headDir, style, true);
                            RenderFeature(btr, tr, headPts, style, allIds);
                        }

                        if (style.StartCapStyle != ArrowHeadType.None)
                        {
                            Vector3d tailDir = path.GetFirstDerivative(path.StartParam).GetNormal().Negate();
                            Point3dCollection tailPts = CalculateFeaturePoints(path.StartPoint, tailDir, style, false);
                            RenderFeature(btr, tr, tailPts, style, allIds);
                        }
                    }

                    // 自动打组：方便用户一键选中/删除整条动线
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
        /// 渲染动线主体：实现 Solid、Dashed 和 Pattern 三类逻辑
        /// </summary>
        private static void RenderBody(BlockTableRecord btr, Transaction tr, Curve curve, AnalysisStyle style, ObjectIdCollection ids)
        {
            double len = curve.GetDistanceAtParameter(curve.EndParam);

            // 分支 1：实线渲染 (Solid)
            if (style.PathType == PathCategory.Solid)
            {
                // ✨ 动态采样逻辑：基于长度和尺度计算采样点数，确保转弯圆滑
                // 采样步长参考箭头大小的 15%，最小不小于 0.5 单位
                double step = Math.Max(style.ArrowSize * 0.15, 0.5);
                int samples = (int)Math.Max(len / step, 100); // 确保最少 100 个采样点

                // 性能保护：防止极长线导致采样点过多（上限 3000）
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
                            // 设置贝塞尔渐变宽度
                            double startW = CalculateBezierWidth(t, style.StartWidth, style.MidWidth, style.EndWidth);
                            double endW = CalculateBezierWidth((i + 1.0) / samples, style.StartWidth, style.MidWidth, style.EndWidth);
                            pl.SetStartWidthAt(i, startW);
                            pl.SetEndWidthAt(i, endW);
                        }
                    }
                    ids.Add(AddToDb(btr, tr, pl, style));
                }
            }
            // 分支 2：阵列模式 (Pattern)
            else if (style.PathType == PathCategory.Pattern)
            {
                ObjectId blockId = GetOrImportBlock(style.SelectedBlockName, btr.Database, tr);
                if (blockId.IsNull) return;

                double dist = 0;
                while (dist <= len)
                {
                    Point3d pt = curve.GetPointAtDist(dist);
                    Vector3d tan = curve.GetFirstDerivative(curve.GetParameterAtDistance(dist)).GetNormal();
                    using (BlockReference br = new BlockReference(pt, blockId))
                    {
                        br.Rotation = Math.Atan2(tan.Y, tan.X);
                        br.ScaleFactors = new Scale3d(style.PatternScale);
                        ids.Add(AddToDb(btr, tr, br, style));
                    }
                    dist += style.PatternSpacing;
                }
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

        private static Point3dCollection CalculateFeaturePoints(Point3d basePt, Vector3d dir, AnalysisStyle style, bool isHead)
        {
            Point3dCollection pts = new Point3dCollection();
            double s = style.ArrowSize;
            Vector3d norm = dir.GetPerpendicularVector();
            ArrowHeadType type = isHead ? style.EndCapStyle : style.StartCapStyle;

            switch (type)
            {
                case ArrowHeadType.SwallowTail:
                    double d = s * style.SwallowDepth;
                    pts.Add(basePt);
                    pts.Add(basePt - dir * s + norm * s * 0.5);
                    pts.Add(basePt - dir * d);
                    pts.Add(basePt - dir * s - norm * s * 0.5);
                    break;
                case ArrowHeadType.Circle:
                    pts.Add(basePt + norm * s * 0.5);
                    pts.Add(basePt - norm * s * 0.5);
                    break;
                case ArrowHeadType.Basic:
                default:
                    pts.Add(basePt);
                    pts.Add(basePt - dir * s + norm * s * 0.4);
                    pts.Add(basePt - dir * s - norm * s * 0.4);
                    break;
            }
            return pts;
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

        private static ObjectId AddToDb(BlockTableRecord btr, Transaction tr, Entity ent, AnalysisStyle style)
        {
            ent.Layer = style.TargetLayer;
            ent.Color = Color.FromRgb(style.MainColor.R, style.MainColor.G, style.MainColor.B);
            ObjectId id = btr.AppendEntity(ent);
            tr.AddNewlyCreatedDBObject(ent, true);
            return id;
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

        private static Curve TrimPath(Curve rawPath, double headLen, double tailLen)
        {
            try
            {
                double totalLen = rawPath.GetDistanceAtParameter(rawPath.EndParam);
                if (totalLen <= (headLen + tailLen + 1e-4)) return null;
                double p1 = rawPath.GetParameterAtDistance(tailLen);
                double p2 = rawPath.GetParameterAtDistance(totalLen - headLen);
                using (DBObjectCollection subCurves = rawPath.GetSplitCurves(new DoubleCollection { p1, p2 }))
                {
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