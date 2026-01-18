using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Plugin_AnalysisMaster.Models;
using System;
using System.Collections.Generic;

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
            if (doc == null) return;

            // ✨ 核心修复：在非模态窗口中修改数据库，必须显式锁定文档
            using (DocumentLock dl = doc.LockDocument())
            {
                Database db = doc.Database;
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    // 之前的报错就发生在这里的 OpenMode.ForWrite
                    BlockTableRecord btr = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                    // ... 后续逻辑保持不变 ...
                    Vector3d dir = endPt - startPt;
                    if (dir.Length < 1e-6) return;
                    Vector3d unitDir = dir.GetNormal();
                    Vector3d normal = unitDir.GetPerpendicularVector();

                    Point3dCollection pts = new Point3dCollection();
                    if (style.HeadType == ArrowHeadType.SwallowTail)
                    {
                        pts = GetSwallowTailPoints(startPt, endPt, unitDir, normal, style);
                    }
                    else
                    {
                        pts = GetBasicArrowPoints(startPt, endPt, unitDir, normal, style);
                    }

                    using (Polyline pl = new Polyline())
                    {
                        for (int i = 0; i < pts.Count; i++)
                        {
                            pl.AddVertexAt(i, new Point2d(pts[i].X, pts[i].Y), 0, 0, 0);
                        }
                        pl.Closed = true;
                        // 注意这里使用的 Color 已经在 AnalysisStyle 中修复为 WPF 类型
                        pl.Color = Autodesk.AutoCAD.Colors.Color.FromRgb(style.MainColor.R, style.MainColor.G, style.MainColor.B);

                        btr.AppendEntity(pl);
                        tr.AddNewlyCreatedDBObject(pl, true);

                        CreateSolidHatch(btr, tr, pl, style);
                    }

                    tr.Commit();
                }
            } // 文档锁在这里自动释放
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

                // ✨ 修正 1：PreDefined 的 D 需要大写
                hatch.SetHatchPattern(HatchPatternType.PreDefined, "SOLID");

                hatch.Color = boundary.Color;

                // ✨ 修正 2：将 double 显式转换为 byte
                hatch.Transparency = new Autodesk.AutoCAD.Colors.Transparency((byte)style.Transparency);

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