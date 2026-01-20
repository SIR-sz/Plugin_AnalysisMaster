using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Runtime;
using CadAtlasManager.Core; // 引用接口定义库
using Plugin_AnalysisMaster.UI; // 引用 UI 命名空间
using System;
using System.Windows.Media;

// [必选] 注册命令类，确保 AutoCAD 命令行能识别 [CommandMethod]
[assembly: CommandClass(typeof(Plugin_AnalysisMaster.MainTool))]

namespace Plugin_AnalysisMaster
{
    /// <summary>
    /// 动态曲线助手插件入口
    /// </summary>
    public class MainTool : ICadTool
    {
        // 授权标记：仅在主程序点击 Execute 或 Debug/Standalone 模式下开启
        private static bool _isAuthorized = false;

        #region --- ICadTool 接口实现 ---

        public string ToolName => "动态曲线助手"; // 显示在主面板上的名称
        public string IconCode => "\uE81C";      // 使用路径图标
        public string Description => "专业级动线、箭头与分析线绘制工具";
        public string Category { get; set; } = "绘图增强";

        // 主程序会自动寻找同名 PNG 并填充此属性
        public ImageSource ToolPreview { get; set; }

        public bool VerifyHost(Guid hostGuid)
        {
            // 严格比对主程序身份暗号
            return hostGuid == new Guid("A7F3E2B1-4D5E-4B8C-9F0A-1C2B3D4E5F6B");
        }

        public void Execute()
        {
            // 通过主程序点击，视为授权成功
            _isAuthorized = true;

            // 启动 UI 界面
            MainControlWindow.ShowTool();
        }

        #endregion

        #region --- 命令行入口 ---

        [CommandMethod("DRAW_ANALYSIS")]
        [CommandMethod("DXFX")] // 简写命令
        public void MainCommandEntry()
        {
#if STANDALONE || DEBUG
            // 调试或独立模式：直接运行
            ShowUIInternal();
#else
            // Release 模式：检查授权
            if (_isAuthorized)
            {
                ShowUIInternal();
            }
            else
            {
                var doc = Application.DocumentManager.MdiActiveDocument;
                doc?.Editor.WriteMessage("\n[错误] 该插件为 智汇CAD全流程管理系统 授权版，请从主程序面板启动。");
            }
#endif
        }

        private void ShowUIInternal()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            doc.Editor.WriteMessage($"\n[{ToolName}] 正在启动界面...");

            // 调用 UI 层的静态启动方法
            MainControlWindow.ShowTool();
        }

        #endregion
    }
}