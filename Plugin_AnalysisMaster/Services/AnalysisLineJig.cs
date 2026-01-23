using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.GraphicsInterface;
using Plugin_AnalysisMaster.Models;
using System;
using Polyline = Autodesk.AutoCAD.DatabaseServices.Polyline;

namespace Plugin_AnalysisMaster.Services
{
    public class AnalysisLineJig : DrawJig
    {
        private Point3dCollection _pts;
        private Point3d _tempPt;
        private AnalysisStyle _style;

        public Point3d LastPoint => _tempPt;

        public AnalysisLineJig(Point3d startPt, AnalysisStyle style)
        {
            _pts = new Point3dCollection { startPt };
            _tempPt = startPt;
            _style = style;
        }

        public Point3dCollection GetPoints() => _pts;

        /// <summary>
        /// 智能采集算法：针对曲线和直线采用不同的精简策略。
        /// </summary>
        public void AddPoint(Point3d pt)
        {
            if (_pts.Count == 0) { _pts.Add(pt); return; }
            if (pt.DistanceTo(_pts[_pts.Count - 1]) < 0.001) return;
            if (_pts.Count < 2) { _pts.Add(pt); return; }

            Vector3d vOld = (_pts[_pts.Count - 1] - _pts[_pts.Count - 2]).GetNormal();
            Vector3d vNew = (pt - _pts[_pts.Count - 1]).GetNormal();
            double dot = vOld.DotProduct(vNew);

            // ✨ 阈值微调：
            // 折线模式 0.9999 (极高精简)
            // 曲线模式改为 0.999 (约 2.5 度) 或 0.9995 (约 1.8 度)，
            // 这会让曲线在接近直线的部分也能够合并夹点。
            double threshold = _style.IsCurved ? 0.9995 : 0.9999;

            if (dot > threshold)
                _pts[_pts.Count - 1] = pt; // ✨ 核心：共线则替换最后一点，不新增
            else
                _pts.Add(pt);

            _tempPt = pt;
        }

        protected override SamplerStatus Sampler(JigPrompts prompts)
        {
            JigPromptPointOptions jppo = new JigPromptPointOptions("\n选择下一个控制点 [回车/右键结束]: ");
            jppo.UserInputControls = UserInputControls.Accept3dCoordinates | UserInputControls.NullResponseAccepted | UserInputControls.AnyBlankTerminatesInput;
            if (_pts.Count > 0) { jppo.BasePoint = _pts[_pts.Count - 1]; jppo.UseBasePoint = true; }

            PromptPointResult ppr = prompts.AcquirePoint(jppo);
            if (ppr.Status != PromptStatus.OK) return SamplerStatus.Cancel;
            if (ppr.Value.DistanceTo(_tempPt) < 0.001) return SamplerStatus.NoChange;

            _tempPt = ppr.Value;
            return SamplerStatus.OK;
        }

        protected override bool WorldDraw(WorldDraw draw)
        {
            Point3dCollection previewPoints = new Point3dCollection();
            foreach (Point3d p in _pts) previewPoints.Add(p);
            previewPoints.Add(_tempPt);

            if (previewPoints.Count < 2) return true;

            try
            {
                Curve path;
                if (_style.IsCurved && previewPoints.Count > 2)
                    path = new Spline(previewPoints, 3, 0);
                else
                {
                    Polyline pl = new Polyline();
                    for (int i = 0; i < previewPoints.Count; i++)
                        pl.AddVertexAt(i, new Point2d(previewPoints[i].X, previewPoints[i].Y), 0, 0, 0);
                    path = pl;
                }

                using (path)
                {
                    // ✨ 修复：显式定义预览所需的几何变量，解决 CS0103 错误
                    double totalLen = path.GetDistanceAtParameter(path.EndParam);
                    Vector3d tangent = path.GetFirstDerivative(path.EndParam).GetNormal();
                    Vector3d normal = tangent.GetPerpendicularVector();
                    Point3d endPt = path.EndPoint;

                    var acColor = Autodesk.AutoCAD.Colors.Color.FromRgb(_style.MainColor.R, _style.MainColor.G, _style.MainColor.B);
                    draw.SubEntityTraits.TrueColor = acColor.EntityColor;
                    draw.SubEntityTraits.Color = 1;

                    double headIndent = (_style.EndArrowType == "None") ? 0 : _style.CapIndent;
                    if (totalLen > headIndent)
                    {
                        double splitParam = path.GetParameterAtDistance(totalLen - headIndent);
                        using (DBObjectCollection curves = path.GetSplitCurves(new DoubleCollection { splitParam }))
                        {
                            if (curves.Count > 0)
                            {
                                Entity body = (Entity)curves[0];
                                body.Color = acColor;
                                draw.Geometry.Draw(body);
                                body.Dispose();
                            }
                        }
                    }

                    if (_style.EndArrowType != "None")
                    {
                        using (Polyline headPl = new Polyline())
                        {
                            Point3dCollection headPts = CalculatePreviewHead(endPt, tangent, normal, _style);
                            for (int i = 0; i < headPts.Count; i++)
                                headPl.AddVertexAt(i, new Point2d(headPts[i].X, headPts[i].Y), 0, 0, 0);
                            headPl.Closed = true;
                            headPl.Color = acColor;
                            draw.Geometry.Draw(headPl);
                        }
                    }
                }
            }
            catch { }
            return true;
        }

        private Point3dCollection CalculatePreviewHead(Point3d end, Vector3d dir, Vector3d norm, AnalysisStyle style)
        {
            Point3dCollection pts = new Point3dCollection();
            double s = style.ArrowSize;
            pts.Add(end);
            pts.Add(end - dir * s + norm * s * 0.4);
            pts.Add(end - dir * s - norm * s * 0.4);
            return pts;
        }
    }
}