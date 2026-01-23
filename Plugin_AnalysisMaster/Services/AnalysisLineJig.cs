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
        /// 采样逻辑：捕捉鼠标移动并更新临时点。
        /// 修改说明：
        /// 1. 增加了对 ppr.Status 的严格判定。只有在 PromptStatus.OK（即用户明确移动或点击）时才更新临时点。
        /// 2. 解决了“连接到原点”的 Bug：当用户按下回车或右键结束（Status 为 None）时，直接返回 Cancel 结束采样，
        ///    不再执行 _tempPt = ppr.Value，从而避免了将鼠标位置错误识别为坐标原点 (0,0,0) 的问题。
        /// </summary>
        protected override SamplerStatus Sampler(JigPrompts prompts)
        {
            JigPromptPointOptions jppo = new JigPromptPointOptions("\n选择下一个控制点 [回车/右键结束]: ");
            jppo.UserInputControls = UserInputControls.Accept3dCoordinates |
                                     UserInputControls.NullResponseAccepted |
                                     UserInputControls.AnyBlankTerminatesInput;

            if (_pts.Count > 0)
            {
                jppo.BasePoint = _pts[_pts.Count - 1];
                jppo.UseBasePoint = true;
            }

            PromptPointResult ppr = prompts.AcquirePoint(jppo);

            // ✨ 核心修复：如果状态不是 OK（如用户按回车结束、右键或取消），直接停止采样，不更新临时点
            if (ppr.Status != PromptStatus.OK)
                return SamplerStatus.Cancel;

            // 距离太近不刷新，防止画面抖动
            if (ppr.Value.DistanceTo(_tempPt) < 0.001)
                return SamplerStatus.NoChange;

            // ✨ 只有在 OK 状态下才更新点位
            _tempPt = ppr.Value;
            return SamplerStatus.OK;
        }

        /// <summary>
        /// 绘图逻辑：在屏幕上实时绘制“假”的图元进行预览。
        /// 修改说明：
        /// 1. 修复了 CS0200 错误：不再尝试对只读的 Color.ColorIndex 赋值。
        /// 2. 采用了正确的 ACI 强制逻辑：通过直接设置 SubEntityTraits.Color 和实体的 ColorIndex 属性来确保预览可见。
        /// 3. 实现双重保险：SubEntityTraits.TrueColor 负责显示用户选择的真彩色，而 Color = 1 (红色) 负责在某些显示模式下提供不为白色的对比色。
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

                    // ✨ 获取真彩色对象
                    var acColor = Autodesk.AutoCAD.Colors.Color.FromRgb(
                        _style.MainColor.R, _style.MainColor.G, _style.MainColor.B);

                    // ✨ 修复位置：在 SubEntityTraits 上分别设置真彩色和索引色
                    // TrueColor 负责显示 UI 选择的颜色，Color 负责强制 ACI 索引为 1 (红色)
                    draw.SubEntityTraits.TrueColor = acColor.EntityColor;
                    draw.SubEntityTraits.Color = 1;

                    // 绘制线体部分
                    double headIndent = (_style.EndArrowType == "None") ? 0 : _style.CapIndent;
                    if (totalLen > headIndent)
                    {
                        double splitParam = path.GetParameterAtDistance(totalLen - headIndent);
                        using (DBObjectCollection curves = path.GetSplitCurves(new DoubleCollection { splitParam }))
                        {
                            if (curves.Count > 0)
                            {
                                Entity body = (Entity)curves[0];

                                // ✨ 修复位置：直接设置实体的颜色属性
                                // 先设为真彩色
                                body.Color = acColor;
                                // 注意：在 AutoCAD 中设置 Entity.Color 会覆盖索引，反之亦然。
                                // 如果设置 body.ColorIndex = 1 会丢失真彩色。
                                // 所以我们依赖上面 draw.SubEntityTraits.Color = 1 的上下文覆盖。

                                draw.Geometry.Draw(body);
                                body.Dispose();
                            }
                        }
                    }

                    // 绘制端头（三角形）预览
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

                            // 设置端头颜色
                            headPl.Color = acColor;

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