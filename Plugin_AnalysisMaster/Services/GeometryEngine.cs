using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.GraphicsInterface;
using Plugin_AnalysisMaster.Models;
using System;
using System.Collections.Generic;
using Polyline = Autodesk.AutoCAD.DatabaseServices.Polyline;



namespace Plugin_AnalysisMaster.Services
{
    public class AnalysisLineJig : DrawJig
    {
        private Point3dCollection _pts;    // 已确定的点
        private Point3d _tempPt;           // 当前鼠标悬停的点
        private AnalysisStyle _style;      // 样式配置
        public Point3d LastPoint => _tempPt;

        public AnalysisLineJig(Point3d startPt, AnalysisStyle style)
        {
            _pts = new Point3dCollection { startPt };
            _tempPt = startPt;
            _style = style;
        }

        public Point3dCollection GetPoints() => _pts;

        // 核心采样：监听鼠标移动
        protected override SamplerStatus Sampler(JigPrompts prompts)
        {
            JigPromptPointOptions jppo = new JigPromptPointOptions("\n选择下一个控制点 [回车/右键结束]: ");
            jppo.UserInputControls = UserInputControls.Accept3dCoordinates |
                                     UserInputControls.NullResponseAccepted |
                                     UserInputControls.AnyBlankTerminatesInput;

            PromptPointResult ppr = prompts.AcquirePoint(jppo);

            if (ppr.Status == PromptStatus.Cancel) return SamplerStatus.Cancel;

            // 距离太近不刷新，防止抖动
            if (ppr.Value.DistanceTo(_tempPt) < 0.001)
                return SamplerStatus.NoChange;

            _tempPt = ppr.Value;
            return SamplerStatus.OK;
        }

        // 核心绘图：实时预览线条和箭头
        protected override bool WorldDraw(WorldDraw draw)
        {
            // 构造包含当前鼠标位置的临时点集
            Point3dCollection previewPoints = new Point3dCollection();
            foreach (Point3d p in _pts) previewPoints.Add(p);
            previewPoints.Add(_tempPt);

            if (previewPoints.Count < 2) return true;

            try
            {
                // 1. 生成临时路径 (Spline 或 Polyline)
                Curve path;
                if (_style.IsCurved && previewPoints.Count > 2)
                    path = new Spline(previewPoints, 3, 0); //
                else
                {
                    Polyline pl = new Polyline();
                    for (int i = 0; i < previewPoints.Count; i++)
                        pl.AddVertexAt(i, new Point2d(previewPoints[i].X, previewPoints[i].Y), 0, 0, 0);
                    path = pl;
                }

                using (path)
                {
                    double totalLen = path.GetDistanceAtParameter(path.EndParam); //
                    Vector3d tangent = path.GetFirstDerivative(path.EndParam).GetNormal();
                    Vector3d normal = tangent.GetPerpendicularVector();
                    Point3d endPt = path.EndPoint;

                    // 设置预览颜色（使用样式中定义的颜色）
                    draw.SubEntityTraits.Color = Autodesk.AutoCAD.Colors.Color.FromRgb(
                        _style.MainColor.R, _style.MainColor.G, _style.MainColor.B).ColorIndex;

                    // 2. 绘制本体线预览（缩短后的，防止重叠）
                    double arrowSize = _style.ArrowSize;
                    if (totalLen > arrowSize)
                    {
                        double splitParam = path.GetParameterAtDistance(totalLen - arrowSize); //
                        using (DBObjectCollection curves = path.GetSplitCurves(new DoubleCollection { splitParam }))
                        {
                            if (curves.Count > 0)
                            {
                                Entity body = (Entity)curves[0];
                                draw.Geometry.Draw(body);
                                body.Dispose();
                            }
                        }
                    }

                    // 3. 绘制箭头头部预览
                    using (Polyline arrowPl = new Polyline())
                    {
                        Point3dCollection arrowPts = _style.HeadType == ArrowHeadType.SwallowTail
                            ? CalculateSwallowTail(endPt, tangent, normal, _style)
                            : CalculateBasicArrow(endPt, tangent, normal, _style);

                        for (int i = 0; i < arrowPts.Count; i++)
                            arrowPl.AddVertexAt(i, new Point2d(arrowPts[i].X, arrowPts[i].Y), 0, 0, 0);
                        arrowPl.Closed = true;

                        draw.Geometry.Draw(arrowPl);
                    }
                }
            }
            catch { }

            return true;
        }

        // 内部复用 GeometryEngine 的坐标算法逻辑
        private Point3dCollection CalculateSwallowTail(Point3d end, Vector3d uDir, Vector3d uNormal, AnalysisStyle style)
        {
            double size = style.ArrowSize;
            double depth = style.ArrowSize * style.SwallowDepth;
            Point3dCollection pts = new Point3dCollection();
            pts.Add(end);
            Point3d backBase = end - (uDir * size);
            pts.Add(backBase + (uNormal * (size * 0.5)));
            pts.Add(backBase + (uDir * depth));
            pts.Add(backBase - (uNormal * (size * 0.5)));
            return pts;
        }

        private Point3dCollection CalculateBasicArrow(Point3d end, Vector3d uDir, Vector3d uNormal, AnalysisStyle style)
        {
            double size = style.ArrowSize;
            Point3dCollection pts = new Point3dCollection();
            pts.Add(end);
            Point3d backBase = end - (uDir * size);
            pts.Add(backBase + (uNormal * (size * 0.4)));
            pts.Add(backBase - (uNormal * (size * 0.4)));
            return pts;
        }
    }
    public class GeometryEngine
    {
        /// <summary>
        /// 绘制分析线的核心方法
        /// </summary>
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

                    // 1. 创建路径曲线
                    Curve mainPath;
                    if (style.IsCurved && points.Count > 2)
                    {
                        // 使用样条曲线实现平滑多点路径
                        mainPath = new Spline(points, 3, 0);
                    }
                    else
                    {
                        Polyline pl = new Polyline();
                        for (int i = 0; i < points.Count; i++)
                            pl.AddVertexAt(i, new Point2d(points[i].X, points[i].Y), 0, 0, 0);
                        mainPath = pl;
                    }

                    using (mainPath)
                    {
                        // 2. 获取参数和切线
                        double endParam = mainPath.EndParam;

                        // ✨ 修正：AutoCAD 2015 中使用全称方法名
                        double totalDist = mainPath.GetDistanceAtParameter(endParam);
                        Vector3d tangent = mainPath.GetFirstDerivative(endParam).GetNormal();
                        Vector3d normal = tangent.GetPerpendicularVector();
                        Point3d endPt = mainPath.EndPoint;

                        // 3. 绘制动线本体（自动缩短以留出箭头空间）
                        double arrowSize = style.ArrowSize;
                        if (totalDist > arrowSize)
                        {
                            double splitDist = totalDist - arrowSize;
                            // ✨ 修正：AutoCAD 2015 中使用全称方法名
                            double splitParam = mainPath.GetParameterAtDistance(splitDist);

                            using (DBObjectCollection subCurves = mainPath.GetSplitCurves(new DoubleCollection { splitParam }))
                            {
                                if (subCurves.Count > 0)
                                {
                                    Curve body = subCurves[0] as Curve;
                                    body.Layer = style.TargetLayer;
                                    body.Color = Autodesk.AutoCAD.Colors.Color.FromRgb(style.MainColor.R, style.MainColor.G, style.MainColor.B);
                                    body.LineWeight = MapLineWeight(style.LineWeight);

                                    btr.AppendEntity(body);
                                    tr.AddNewlyCreatedDBObject(body, true);
                                }
                            }
                        }

                        // 4. 绘制箭头（基于切线方向计算，确保与曲线末端精准对齐）
                        Point3dCollection arrowPts = style.HeadType == ArrowHeadType.SwallowTail
                            ? GetSwallowTailPoints(Point3d.Origin, endPt, tangent, normal, style)
                            : GetBasicArrowPoints(Point3d.Origin, endPt, tangent, normal, style);

                        using (Polyline headPl = new Polyline())
                        {
                            for (int i = 0; i < arrowPts.Count; i++)
                                headPl.AddVertexAt(i, new Point2d(arrowPts[i].X, arrowPts[i].Y), 0, 0, 0);

                            headPl.Closed = true;
                            headPl.Layer = style.TargetLayer;
                            headPl.Color = Autodesk.AutoCAD.Colors.Color.FromRgb(style.MainColor.R, style.MainColor.G, style.MainColor.B);

                            btr.AppendEntity(headPl);
                            tr.AddNewlyCreatedDBObject(headPl, true);
                            CreateSolidHatch(btr, tr, headPl, style);
                        }
                    }
                    tr.Commit();
                }
            }
        }
        // ✨ 将 double (0.0-2.11) 映射到 AutoCAD 枚举
        private static LineWeight MapLineWeight(double weight)
        {
            if (weight <= 0.0) return LineWeight.LineWeight000;
            if (weight <= 0.15) return LineWeight.LineWeight015;
            if (weight <= 0.30) return LineWeight.LineWeight030;
            if (weight <= 0.50) return LineWeight.LineWeight050;
            return LineWeight.LineWeight080;
        }
        // ✨ 新增：图层管理逻辑
        private static void EnsureLayer(Database db, Transaction tr, AnalysisStyle style)
        {
            LayerTable lt = tr.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable;
            if (!lt.Has(style.TargetLayer))
            {
                lt.UpgradeOpen();
                using (LayerTableRecord ltr = new LayerTableRecord())
                {
                    ltr.Name = style.TargetLayer;
                    ltr.Color = Autodesk.AutoCAD.Colors.Color.FromRgb(style.MainColor.R, style.MainColor.G, style.MainColor.B);
                    lt.Add(ltr);
                    tr.AddNewlyCreatedDBObject(ltr, true);
                }
            }
        }
        /// <summary>
        /// 计算燕尾箭头的 5 个顶点坐标
        /// </summary>
        // 箭头坐标算法：基于传入的 tangent (切线) 重新计算，不再依赖 start-end 直线方向
        // 箭头坐标算法：根据传入的切线向量 (uDir) 重新计算
        private static Point3dCollection GetSwallowTailPoints(Point3d unused, Point3d end, Vector3d uDir, Vector3d uNormal, AnalysisStyle style)
        {
            double size = style.ArrowSize;
            double depth = style.ArrowSize * style.SwallowDepth;
            Point3dCollection pts = new Point3dCollection();
            pts.Add(end);
            Point3d backBase = end - (uDir * size);
            pts.Add(backBase + (uNormal * (size * 0.5)));
            pts.Add(backBase + (uDir * depth));
            pts.Add(backBase - (uNormal * (size * 0.5)));
            return pts;
        }

        private static void CreateSolidHatch(BlockTableRecord btr, Transaction tr, Polyline boundary, AnalysisStyle style)
        {
            using (Hatch hatch = new Hatch())
            {
                btr.AppendEntity(hatch);
                tr.AddNewlyCreatedDBObject(hatch, true);
                hatch.SetHatchPattern(HatchPatternType.PreDefined, "SOLID");
                hatch.Color = boundary.Color;
                hatch.Transparency = new Autodesk.AutoCAD.Colors.Transparency((byte)style.Transparency);
                hatch.Layer = boundary.Layer;
                ObjectIdCollection ids = new ObjectIdCollection { boundary.ObjectId };
                hatch.AppendLoop(HatchLoopTypes.Default, ids);
                hatch.EvaluateHatch(true);
            }
        }

        private static Point3dCollection GetBasicArrowPoints(Point3d unused, Point3d end, Vector3d uDir, Vector3d uNormal, AnalysisStyle style)
        {
            double size = style.ArrowSize;
            Point3dCollection pts = new Point3dCollection();
            pts.Add(end);
            Point3d backBase = end - (uDir * size);
            pts.Add(backBase + (uNormal * (size * 0.4)));
            pts.Add(backBase - (uNormal * (size * 0.4)));
            return pts;
        }
    }
}