using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Runtime;
using CadAtlasManager.Core;
using Plugin_AnalysisMaster.UI;
using System;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using System.Windows.Media;
using VMProtect;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: ExtensionApplication(typeof(Plugin_AnalysisMaster.MainTool))]
[assembly: CommandClass(typeof(Plugin_AnalysisMaster.MainTool))]

namespace Plugin_AnalysisMaster
{
    public class MainTool : ICadTool, IExtensionApplication
    {
        private static AnimationWindow _animWindow = null;
        private static System.Windows.Forms.Timer _heartbeatTimer = null;

        #region --- IExtensionApplication 接口实现 (改为静默加载) ---

        public void Initialize()
        {
            // ✨ 移除此处 CheckLicense 调用，实现 CAD 启动时不弹出激活窗体
        }

        public void Terminate()
        {
            StopHeartbeat();
        }

        #endregion

        #region --- ICadTool 接口实现 ---

        public string ToolName => "动态曲线助手";
        public string IconCode => "\uE81C";
        public string Description => "专业级动线、箭头与分析线绘制工具";
        public string Category { get; set; } = "绘图增强";
        public ImageSource ToolPreview { get; set; }

        public bool VerifyHost(Guid hostGuid)
        {
            return hostGuid == new Guid("A7F3E2B1-4D5E-4B8C-9F0A-1C2B3D4E5F6B");
        }

        public void Execute()
        {
            // ✨ 第一次点击执行时检测授权
            if (!CheckLicense()) return;
            ShowUIInternal();
        }

        #endregion

        #region --- 授权验证核心逻辑 (参考 CadAtlasManager 布局) ---

        [VMProtect.Begin]
        public static bool CheckLicense()
        {
            int status = (int)VMProtect.SDK.GetSerialNumberState();
            if (status == 0) return true;

            try
            {
                string dllPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                string licPath = Path.Combine(dllPath, "key.lic");
                if (File.Exists(licPath))
                {
                    string cleanKey = File.ReadAllText(licPath).Trim();
                    if ((int)VMProtect.SDK.SetSerialNumber(cleanKey) == 0) return true;
                }
            }
            catch { }

            string hwid = VMProtect.SDK.GetCurrentHWID();
            string msg = (status == 3) ? "授权时间/单次运行时间已到期。" : "未检测到有效授权，请激活后使用。";

            return ShowActivationDialog(msg, hwid);
        }

        [VMProtect.Begin]
        private static bool ShowActivationDialog(string message, string hwid)
        {
            bool isSuccess = false;
            using (Form form = new Form())
            {
                // ✨ 严格参考 CadAtlasManager 的窗体参数
                form.Text = "动态曲线助手 - 软件激活";
                form.Size = new System.Drawing.Size(480, 450);
                form.StartPosition = FormStartPosition.CenterScreen;
                form.FormBorderStyle = FormBorderStyle.FixedDialog;
                form.MaximizeBox = false; form.MinimizeBox = false;
                form.TopMost = true;

                Label lblMsg = new Label() { Text = message + "\n请输入注册码激活系统：", Dock = DockStyle.Top, Height = 60, Padding = new Padding(10), ForeColor = System.Drawing.Color.Red, Font = new System.Drawing.Font("微软雅黑", 9F) };

                // ✨ 修复复制按钮变形：使用固定高度并分层 Dock
                GroupBox grpHwid = new GroupBox() { Text = "1. 复制机器码发给管理员", Dock = DockStyle.Top, Height = 80, Padding = new Padding(5) };
                System.Windows.Forms.TextBox txtHwid = new System.Windows.Forms.TextBox() { Text = hwid, Dock = DockStyle.Top, ReadOnly = true, Font = new System.Drawing.Font("Consolas", 10F) };
                System.Windows.Forms.Button btnCopy = new System.Windows.Forms.Button() { Text = "点击复制机器码", Dock = DockStyle.Bottom, Height = 28 };
                btnCopy.Click += (s, e) => { Clipboard.SetText(hwid); MessageBox.Show("机器码已复制到剪贴板！"); };
                grpHwid.Controls.Add(btnCopy); grpHwid.Controls.Add(txtHwid);

                GroupBox grpInput = new GroupBox() { Text = "2. 输入注册码", Dock = DockStyle.Top, Height = 180, Padding = new Padding(5) };
                System.Windows.Forms.TextBox txtKey = new System.Windows.Forms.TextBox() { Multiline = true, Dock = DockStyle.Fill, ScrollBars = ScrollBars.Vertical, Font = new System.Drawing.Font("Consolas", 9F) };
                grpInput.Controls.Add(txtKey);

                System.Windows.Forms.Button btnActivate = new System.Windows.Forms.Button() { Text = "立即激活系统", Dock = DockStyle.Bottom, Height = 50, Font = new System.Drawing.Font("微软雅黑", 10F, System.Drawing.FontStyle.Bold) };
                btnActivate.Click += (s, e) =>
                {
                    string keyInput = txtKey.Text.Trim();
                    if ((int)VMProtect.SDK.SetSerialNumber(keyInput) == 0)
                    {
                        try { File.WriteAllText(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "key.lic"), keyInput); } catch { }
                        MessageBox.Show("激活成功！");
                        isSuccess = true;
                        form.Close();
                    }
                    else MessageBox.Show("注册码无效！");
                };

                form.Controls.Add(grpInput);
                form.Controls.Add(grpHwid);
                form.Controls.Add(lblMsg);
                form.Controls.Add(btnActivate);
                form.ShowDialog();
            }
            return isSuccess;
        }

        private static void StartHeartbeat()
        {
            if (_heartbeatTimer != null) return;
            _heartbeatTimer = new System.Windows.Forms.Timer { Interval = 60000 }; // 1 分钟心跳
            _heartbeatTimer.Tick += (s, e) =>
            {
                if (VMProtect.SDK.GetSerialNumberState() != 0)
                {
                    StopHeartbeat();
                    MainControlWindow.CloseTool();
                    CheckLicense();
                }
            };
            _heartbeatTimer.Start();
        }

        private static void StopHeartbeat()
        {
            _heartbeatTimer?.Stop();
            _heartbeatTimer?.Dispose();
            _heartbeatTimer = null;
        }

        #endregion

        #region --- 启动逻辑 ---

        private static void ShowUIInternal()
        {
            MainControlWindow.ShowTool();
            StartHeartbeat();
        }

        [CommandMethod("DRAW_ANALYSIS")]
        [CommandMethod("DXFX")]
        public void MainCommandEntry()
        {
            // ✨ 输入命令时才弹出激活界面
            if (CheckLicense())
            {
                ShowUIInternal();
            }
        }

        #endregion
    }
}