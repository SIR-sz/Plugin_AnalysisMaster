using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.GraphicsInterface;
using Plugin_AnalysisMaster.Models;
using System;
// 解决 2015 版本命名空间冲突
using Polyline = Autodesk.AutoCAD.DatabaseServices.Polyline;

namespace Plugin_AnalysisMaster.Services
{
    /// <summary>
    /// 动线动态预览引擎
    /// </summary>
    public class AnalysisLineJig : DrawJig
    {
        private Point3dCollection _pts;    // 已确定的控制点
        private Point3d _tempPt;           // 当前鼠标实时位置
        private AnalysisStyle _style;      // 样式参数

        // ✨ 供 UI 层访问：获取用户点击左键那一刻，鼠标所在的确切位置
        public Point3d LastPoint => _tempPt;

        public AnalysisLineJig(Point3d startPt, AnalysisStyle style)
        {
            _pts = new Point3dCollection { startPt };
            _tempPt = startPt;
            _style = style;
        }

        // ✨ 供 UI 层访问：获取当前已记录的所有点
        public Point3dCollection GetPoints() => _pts;

        /// <summary>
        /// 采样逻辑：捕捉鼠标移动并更新临时点
        /// </summary>
        protected override SamplerStatus Sampler(JigPrompts prompts)
        {
            JigPromptPointOptions jppo = new JigPromptPointOptions("\n选择下一个控制点 [回车/右键结束]: ");
            jppo.UserInputControls = UserInputControls.Accept3dCoordinates |
                                     UserInputControls.NullResponseAccepted |
                                     UserInputControls.AnyBlankTerminatesInput;

            PromptPointResult ppr = prompts.AcquirePoint(jppo);

            if (ppr.Status == PromptStatus.Cancel) return SamplerStatus.Cancel;

            // 距离太近不刷新，防止画面抖动
            if (ppr.Value.DistanceTo(_tempPt) < 0.001)
                return SamplerStatus.NoChange;

            _tempPt = ppr.Value;
            return SamplerStatus.OK;
        }

        /// <summary>
        /// 绘图逻辑：在屏幕上实时绘制“假”的图元进行预览
        /// </summary>
        protected override bool WorldDraw(WorldDraw draw)
        {
            // 构造包含当前鼠标位置的临时预览点集
            Point3dCollection previewPoints = new Point3dCollection();
            foreach (Point3d p in _pts) previewPoints.Add(p);
            previewPoints.Add(_tempPt);

            if (previewPoints.Count < 2) return true;

            try
            {
                // 1. 生成临时路径 (Spline 或 Polyline)
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
                    double totalLen = path.GetDistanceAtParameter(path.EndParam);
                    // 获取终点切线方向，确保预览时箭头也是对齐的
                    Vector3d tangent = path.GetFirstDerivative(path.EndParam).GetNormal();
                    Vector3d normal = tangent.GetPerpendicularVector();
                    Point3d endPt = path.EndPoint;

                    // 设置预览颜色
                    draw.SubEntityTraits.Color = Autodesk.AutoCAD.Colors.Color.FromRgb(
                        _style.MainColor.R, _style.MainColor.G, _style.MainColor.B).ColorIndex;

                    // 2. 绘制缩短后的本体预览（为箭头留出空间）
                    double headIndent = (_style.HeadType == ArrowHeadType.None) ? 0 : _style.ArrowSize;
                    if (totalLen > headIndent)
                    {
                        double splitParam = path.GetParameterAtDistance(totalLen - headIndent);
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

                    // 3. 绘制端头预览
                    if (_style.HeadType != ArrowHeadType.None)
                    {
                        using (Polyline headPl = new Polyline())
                        {
                            Point3dCollection headPts = CalculatePreviewHead(endPt, tangent, normal, _style);
                            for (int i = 0; i < headPts.Count; i++)
                            {
                                // 圆形端头预览使用凸度处理
                                double bulge = (_style.HeadType == ArrowHeadType.Circle) ? 1.0 : 0.0;
                                headPl.AddVertexAt(i, new Point2d(headPts[i].X, headPts[i].Y), bulge, 0, 0);
                            }
                            headPl.Closed = true;
                            draw.Geometry.Draw(headPl);
                        }
                    }
                }
            }
            catch { }

            return true;
        }

        /// <summary>
        /// 内部简易几何计算：用于预览
        /// </summary>
        private Point3dCollection CalculatePreviewHead(Point3d end, Vector3d dir, Vector3d norm, AnalysisStyle style)
        {
            Point3dCollection pts = new Point3dCollection();
            double s = style.ArrowSize;

            switch (style.HeadType)
            {
                case ArrowHeadType.SwallowTail:
                    double d = s * style.SwallowDepth;
                    pts.Add(end);
                    pts.Add(end - dir * s + norm * s * 0.5);
                    pts.Add(end - dir * d);
                    pts.Add(end - dir * s - norm * s * 0.5);
                    break;
                case ArrowHeadType.Circle:
                    Point3d center = end - dir * s * 0.5;
                    pts.Add(center + norm * s * 0.5);
                    pts.Add(center - norm * s * 0.5);
                    break;
                default: // Basic
                    pts.Add(end);
                    pts.Add(end - dir * s + norm * s * 0.4);
                    pts.Add(end - dir * s - norm * s * 0.4);
                    break;
            }
            return pts;
        }
    }
}