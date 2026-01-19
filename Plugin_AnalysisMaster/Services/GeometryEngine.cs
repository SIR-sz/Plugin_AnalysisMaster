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
        public static void DrawAnalysisLine(Point3dCollection inputPoints, AnalysisStyle style)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Point3dCollection points = inputPoints;

            if (points == null)
            {
                PromptPointOptions ppo = new PromptPointOptions("\n指定起点: ");
                PromptPointResult ppr = doc.Editor.GetPoint(ppo);
                if (ppr.Status != PromptStatus.OK) return;

                // ✨ 修复 CS7036：按正确签名 (Point3d, AnalysisStyle) 实例化 Jig
                var jig = new AnalysisLineJig(ppr.Value, style);
                if (doc.Editor.Drag(jig).Status != PromptStatus.OK) return;
                points = jig.GetPoints();
                points.Add(jig.LastPoint);
            }

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
                        double headIndent = (style.EndCapStyle == ArrowHeadType.None) ? 0 : style.ArrowSize * 0.8;
                        Curve trimmedPath = TrimPath(path, headIndent, 0);
                        if (trimmedPath != null)
                        {
                            RenderBody(btr, tr, trimmedPath, style, allIds);
                            if (trimmedPath != path) trimmedPath.Dispose();
                        }
                    }

                    if (allIds.Count > 1)
                    {
                        DBDictionary gd = (DBDictionary)tr.GetObject(db.GroupDictionaryId, OpenMode.ForWrite);
                        Group grp = new Group("分析动线", true);
                        gd.SetAt("*", grp);
                        tr.AddNewlyCreatedDBObject(grp, true);
                        grp.Append(allIds);
                    }
                    tr.Commit();
                }
            }
        }

        private static void RenderBody(BlockTableRecord btr, Transaction tr, Curve curve, AnalysisStyle style, ObjectIdCollection ids)
        {
            double len = curve.GetDistanceAtParameter(curve.EndParam);
            if (style.PathType == PathCategory.Solid)
            {
                // 实线渲染逻辑...
            }
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

        private static ObjectId GetOrImportBlock(string name, Database destDb, Transaction tr)
        {
            BlockTable bt = (BlockTable)tr.GetObject(destDb.BlockTableId, OpenMode.ForRead);
            if (bt.Has(name)) return bt[name];

            string libPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Assets", "PatternLibrary.dwg");
            if (!File.Exists(libPath)) return ObjectId.Null;

            using (Database srcDb = new Database(false, true))
            {
                // ✨ 修复 CS0117：针对 AutoCAD 2015 优化
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

        private static Curve CreatePathCurve(Point3dCollection pts, bool curved) => (curved && pts.Count > 2) ? (Curve)new Spline(pts, 3, 0) : new Polyline();
        private static Curve TrimPath(Curve c, double h, double t) { try { double l = c.GetDistanceAtParameter(c.EndParam); return c.GetSplitCurves(new DoubleCollection { c.GetParameterAtDistance(t), c.GetParameterAtDistance(l - h) })[1] as Curve; } catch { return null; } }
        private static void EnsureLayer(Database db, Transaction tr, AnalysisStyle s) { /* 逻辑... */ }
    }
}