using System;
using System.Collections.Generic;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.ApplicationServices;
using Plugin_AnalysisMaster.Models;

namespace Plugin_AnalysisMaster.Services
{
    public class GeometryEngine
    {
        /// <summary>
        /// 绘制分析线的核心方法
        /// </summary>
        /// <param name="startPt">起点</param>
        /// <param name="endPt">终点</param>
        /// <param name="style">样式配置模型</param>
        public static void DrawAnalysisLine(Point3d startPt, Point3d endPt, AnalysisStyle style)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord btr = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                // 1. 计算方向向量
                Vector3d dir = endPt - startPt;
                if (dir.Length < 1e-6) return; // 长度太短不处理
                Vector3d unitDir = dir.GetNormal();
                Vector3d normal = unitDir.GetPerpendicularVector(); // 垂直向量，用于计算宽度

                // 2. 根据类型生成不同的顶点
                Point3dCollection pts = new Point3dCollection();

                if (style.HeadType == ArrowHeadType.SwallowTail)
                {
                    pts = GetSwallowTailPoints(startPt, endPt, unitDir, normal, style);
                }
                else
                {
                    // 这里后续可以扩展普通三角形箭头等
                    pts = GetBasicArrowPoints(startPt, endPt, unitDir, normal, style);
                }

                // 3. 创建多段线并填充
                using (Polyline pl = new Polyline())
                {
                    for (int i = 0; i < pts.Count; i++)
                    {
                        pl.AddVertexAt(i, new Point2d(pts[i].X, pts[i].Y), 0, 0, 0);
                    }
                    pl.Closed = true;
                    pl.Color = Autodesk.AutoCAD.Colors.Color.FromRgb(style.MainColor.R, style.MainColor.G, style.MainColor.B);

                    btr.AppendEntity(pl);
                    tr.AddNewlyCreatedDBObject(pl, true);

                    // 如果需要实心填充
                    CreateSolidHatch(btr, tr, pl, style);
                }

                tr.Commit();
            }
        }

        /// <summary>
        /// 计算燕尾箭头的 5 个顶点坐标
        /// </summary>
        private static Point3dCollection GetSwallowTailPoints(Point3d start, Point3d end, Vector3d uDir, Vector3d uNormal, AnalysisStyle style)
        {
            double size = style.ArrowSize;
            double depth = style.ArrowSize * style.SwallowDepth; // 燕尾凹陷深度

            Point3dCollection pts = new Point3dCollection();

            // 顶点 1: 箭头尖端 (End Point)
            pts.Add(end);

            // 计算箭头后端两侧的基准点
            Point3d backBase = end - (uDir * size);

            // 顶点 2: 侧翼 A
            pts.Add(backBase + (uNormal * (size * 0.5)));

            // 顶点 3: 燕尾中心凹陷点
            pts.Add(backBase + (uDir * depth));

            // 顶点 4: 侧翼 B
            pts.Add(backBase - (uNormal * (size * 0.5)));

            return pts;
        }

        private static void CreateSolidHatch(BlockTableRecord btr, Transaction tr, Polyline boundary, AnalysisStyle style)
        {
            using (Hatch hatch = new Hatch())
            {
                btr.AppendEntity(hatch);
                tr.AddNewlyCreatedDBObject(hatch, true);

                hatch.SetHatchPattern(HatchPatternType.Predefined, "SOLID");
                hatch.Color = boundary.Color;
                hatch.Transparency = new Autodesk.AutoCAD.Colors.Transparency(style.Transparency);

                ObjectIdCollection ids = new ObjectIdCollection();
                ids.Add(boundary.ObjectId);
                hatch.AppendLoop(HatchLoopTypes.Default, ids);
                hatch.EvaluateHatch(true);
            }
        }

        private static Point3dCollection GetBasicArrowPoints(Point3d start, Point3d end, Vector3d uDir, Vector3d uNormal, AnalysisStyle style)
        {
            // 普通三角形箭头的逻辑...
            Point3dCollection pts = new Point3dCollection();
            double size = style.ArrowSize;
            pts.Add(end);
            Point3d backBase = end - (uDir * size);
            pts.Add(backBase + (uNormal * (size * 0.4)));
            pts.Add(backBase - (uNormal * (size * 0.4)));
            return pts;
        }
    }
}