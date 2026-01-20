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
        /// 绘图逻辑：在屏幕上实时绘制“假”的图元进行预览。
        /// 修改逻辑：移除对已删除属性 HeadType 的引用。
        /// 如果用户选择了端头图块（EndArrowType != "None"），则在预览中显示一个简易箭头占位。
        /// </summary>
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
                    double totalLen = path.GetDistanceAtParameter(path.EndParam);
                    Vector3d tangent = path.GetFirstDerivative(path.EndParam).GetNormal();
                    Vector3d normal = tangent.GetPerpendicularVector();
                    Point3d endPt = path.EndPoint;

                    draw.SubEntityTraits.Color = Autodesk.AutoCAD.Colors.Color.FromRgb(
                        _style.MainColor.R, _style.MainColor.G, _style.MainColor.B).ColorIndex;

                    // ✨ 核心修改：使用新的属性 CapIndent 进行线体缩进预览
                    double headIndent = (_style.EndArrowType == "None") ? 0 : _style.CapIndent;
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

                    // ✨ 核心修改：如果设置了端头，则显示简易三角形预览
                    if (_style.EndArrowType != "None")
                    {
                        using (Polyline headPl = new Polyline())
                        {
                            Point3dCollection headPts = CalculatePreviewHead(endPt, tangent, normal, _style);
                            for (int i = 0; i < headPts.Count; i++)
                            {
                                headPl.AddVertexAt(i, new Point2d(headPts[i].X, headPts[i].Y), 0, 0, 0);
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
        /// 内部简易几何计算：用于在 Jig 过程中渲染一个简易的三角形箭头预览。
        /// 修改逻辑：删除对旧属性 SwallowDepth 和 HeadType 的引用，统一使用简易三角形。
        /// </summary>
        private Point3dCollection CalculatePreviewHead(Point3d end, Vector3d dir, Vector3d norm, AnalysisStyle style)
        {
            Point3dCollection pts = new Point3dCollection();
            // 使用 ArrowSize 作为预览大小参考
            double s = style.ArrowSize;

            // 默认绘制一个基础的三角形箭头作为占位预览
            pts.Add(end);
            pts.Add(end - dir * s + norm * s * 0.4);
            pts.Add(end - dir * s - norm * s * 0.4);

            return pts;
        }
    }
}