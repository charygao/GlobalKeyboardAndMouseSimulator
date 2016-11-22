using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using System.Xml;
using Microsoft.Win32;
using MouseKeyboardLibrary;

namespace GlobalMacroRecorder
{
    public class MacroForm : Form
    {
        #region Public Fields

        public static BackgroundWorker BackgroundWorker = new BackgroundWorker();

        #endregion Public Fields

        #region Public Constructors


        public MacroForm()
        {
            CheckForIllegalCrossThreadCalls = false;//移除夸UI线程警告

            _macroEventsList = new List<MacroEvent>();

            _tickCount = 0;

            _isPlayRecScript = false;

            //全局鼠标键盘设置
            _globalMouseHook = new MouseHook();
            _globalKeyboardHook = new KeyboardHook();

            //快捷键设置
            _shortcutKeyboardHook = new KeyboardHook();//


            _toolStatus = 0;//0:一般状态 
            icontainer_0 = null;

            InitializeComponent();
            _globalMouseHook.MouseMove += GlobalMouseHookGlobalMouseMove;
            _globalMouseHook.MouseWheel += GlobalMouseHookGlobalMouseWheel;
            _globalMouseHook.MouseDown += GlobalMouseHookGlobalMouseDown;
            _globalMouseHook.MouseUp += GlobalMouseHookGlobalMouseUp;
            _globalKeyboardHook.KeyDown += GlobalKeyboardHookKeyDown;
            _globalKeyboardHook.KeyUp += GlobalKeyboardHookKeyUp;
            _shortcutKeyboardHook.KeyDown += ShortcutKeyboardHookKeyDown;
            _shortcutKeyboardHook.KeyUp += ShortcutKeyboardHookKeyUp;
            BackgroundWorker.DoWork += background_worker;
            BackgroundWorker.WorkerSupportsCancellation = true;
            _keysStartRec = Keys.F9 | Keys.Control;
            _keysStopRec = Keys.F10 | Keys.Control;
            _keysStartPlay = Keys.F11 | Keys.Control;
            _keysStopPlay = Keys.F12 | Keys.Control;
            _shortcutKeyboardHook.Start();
        }

        #endregion Public Constructors

        #region Protected Methods

        protected override void Dispose(bool disposing)
        {
            if (disposing && icontainer_0 != null)
            {
                icontainer_0.Dispose();
            }
            base.Dispose(disposing);
        }

        #endregion Protected Methods

        #region Private Fields

        private bool _isPlayRecScript;

        private Button buttonSetShotcut;
        private IContainer components;
        private ContextMenuStrip contextMenuStrip1;
        private IContainer icontainer_0;
        private int _tickCount;
        private int _toolStatus;
        private KeyboardHook _globalKeyboardHook;
        private KeyboardHook _shortcutKeyboardHook;
        private Keys _keysStartRec;
        private Keys _keysStopRec;
        private Keys _keysStartPlay;
        private Keys _keysStopPlay;
        private Label label1;
        private Label label2;
        private Label label3;
        private Label label4;
        private Label label5;
        private Label label6;
        private List<MacroEvent> _macroEventsList;
        private NumericUpDown loopnumericUpDown;
        private MouseHook _globalMouseHook;
        private NotifyIcon notifyIcon_0;
        private Button openbutton;
        private Button playBackMacroButton;
        private Button recordStartButton;

        private Button recordStopButton;
        private Button savebutton;
        private NumericUpDown spannumericUpDown;
        private Button StopBackMacroButton;
        private TextBox textBoxPlay;
        private TextBox textBoxStart;

        private TextBox textBoxStop;
        private TextBox textBoxStopPlay;
        private ToolStripMenuItem 打开主窗体ToolStripMenuItem;

        private ToolStripMenuItem 访问官网ToolStripMenuItem;

        private ToolStripMenuItem 退出ToolStripMenuItem;

        #endregion Private Fields

        #region Private Methods

        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MacroForm));
            this.recordStartButton = new System.Windows.Forms.Button();
            this.recordStopButton = new System.Windows.Forms.Button();
            this.playBackMacroButton = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.loopnumericUpDown = new System.Windows.Forms.NumericUpDown();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.spannumericUpDown = new System.Windows.Forms.NumericUpDown();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.StopBackMacroButton = new System.Windows.Forms.Button();
            this.savebutton = new System.Windows.Forms.Button();
            this.openbutton = new System.Windows.Forms.Button();
            this.textBoxStart = new System.Windows.Forms.TextBox();
            this.textBoxStop = new System.Windows.Forms.TextBox();
            this.textBoxPlay = new System.Windows.Forms.TextBox();
            this.textBoxStopPlay = new System.Windows.Forms.TextBox();
            this.notifyIcon_0 = new System.Windows.Forms.NotifyIcon(this.components);
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.打开主窗体ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.访问官网ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.退出ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.buttonSetShotcut = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.loopnumericUpDown)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.spannumericUpDown)).BeginInit();
            this.contextMenuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // recordStartButton
            // 
            this.recordStartButton.Location = new System.Drawing.Point(52, 15);
            this.recordStartButton.Name = "recordStartButton";
            this.recordStartButton.Size = new System.Drawing.Size(140, 21);
            this.recordStartButton.TabIndex = 0;
            this.recordStartButton.Text = "开始录制(Ctrl+F9)";
            this.recordStartButton.UseVisualStyleBackColor = true;
            // 
            // recordStopButton
            // 
            this.recordStopButton.Location = new System.Drawing.Point(211, 15);
            this.recordStopButton.Name = "recordStopButton";
            this.recordStopButton.Size = new System.Drawing.Size(143, 21);
            this.recordStopButton.TabIndex = 0;
            this.recordStopButton.Text = "停止录制(Ctrl+F10)";
            this.recordStopButton.UseVisualStyleBackColor = true;
            // 
            // playBackMacroButton
            // 
            this.playBackMacroButton.Location = new System.Drawing.Point(52, 65);
            this.playBackMacroButton.Name = "playBackMacroButton";
            this.playBackMacroButton.Size = new System.Drawing.Size(140, 21);
            this.playBackMacroButton.TabIndex = 1;
            this.playBackMacroButton.Text = "开始播放(Ctrl+F11)";
            this.playBackMacroButton.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(17, 19);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(29, 12);
            this.label1.TabIndex = 2;
            this.label1.Text = "录制";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(17, 69);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(29, 12);
            this.label2.TabIndex = 2;
            this.label2.Text = "播放";
            // 
            // loopnumericUpDown
            // 
            this.loopnumericUpDown.Location = new System.Drawing.Point(429, 65);
            this.loopnumericUpDown.Maximum = new decimal(new int[] {
            -159383553,
            46653770,
            5421,
            0});
            this.loopnumericUpDown.Name = "loopnumericUpDown";
            this.loopnumericUpDown.Size = new System.Drawing.Size(40, 21);
            this.loopnumericUpDown.TabIndex = 3;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(371, 69);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(53, 12);
            this.label3.TabIndex = 4;
            this.label3.Text = "重复次数";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(371, 24);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(53, 12);
            this.label4.TabIndex = 5;
            this.label4.Text = "播放间隔";
            // 
            // spannumericUpDown
            // 
            this.spannumericUpDown.Location = new System.Drawing.Point(429, 20);
            this.spannumericUpDown.Maximum = new decimal(new int[] {
            -1593835521,
            466537709,
            54210,
            0});
            this.spannumericUpDown.Name = "spannumericUpDown";
            this.spannumericUpDown.Size = new System.Drawing.Size(40, 21);
            this.spannumericUpDown.TabIndex = 6;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(474, 24);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(17, 12);
            this.label5.TabIndex = 7;
            this.label5.Text = "秒";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(474, 69);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(17, 12);
            this.label6.TabIndex = 8;
            this.label6.Text = "次";
            // 
            // StopBackMacroButton
            // 
            this.StopBackMacroButton.Location = new System.Drawing.Point(211, 65);
            this.StopBackMacroButton.Name = "StopBackMacroButton";
            this.StopBackMacroButton.Size = new System.Drawing.Size(143, 21);
            this.StopBackMacroButton.TabIndex = 9;
            this.StopBackMacroButton.Text = "停止播放(Ctrl+F12)";
            this.StopBackMacroButton.UseVisualStyleBackColor = true;
            // 
            // savebutton
            // 
            this.savebutton.Location = new System.Drawing.Point(52, 115);
            this.savebutton.Name = "savebutton";
            this.savebutton.Size = new System.Drawing.Size(140, 23);
            this.savebutton.TabIndex = 10;
            this.savebutton.Text = "保存脚本";
            this.savebutton.UseVisualStyleBackColor = true;
            // 
            // openbutton
            // 
            this.openbutton.Location = new System.Drawing.Point(211, 115);
            this.openbutton.Name = "openbutton";
            this.openbutton.Size = new System.Drawing.Size(143, 23);
            this.openbutton.TabIndex = 11;
            this.openbutton.Text = "打开脚本";
            this.openbutton.UseVisualStyleBackColor = true;
            // 
            // textBoxStart
            // 
            this.textBoxStart.Location = new System.Drawing.Point(52, 41);
            this.textBoxStart.Name = "textBoxStart";
            this.textBoxStart.Size = new System.Drawing.Size(140, 21);
            this.textBoxStart.TabIndex = 12;
            this.textBoxStart.Text = "Ctrl+F9";
            this.textBoxStart.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBoxStop
            // 
            this.textBoxStop.Location = new System.Drawing.Point(211, 42);
            this.textBoxStop.Name = "textBoxStop";
            this.textBoxStop.Size = new System.Drawing.Size(143, 21);
            this.textBoxStop.TabIndex = 13;
            this.textBoxStop.Text = "Ctrl+F10";
            this.textBoxStop.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBoxPlay
            // 
            this.textBoxPlay.Location = new System.Drawing.Point(52, 88);
            this.textBoxPlay.Name = "textBoxPlay";
            this.textBoxPlay.Size = new System.Drawing.Size(140, 21);
            this.textBoxPlay.TabIndex = 14;
            this.textBoxPlay.Text = "Ctrl+F11";
            this.textBoxPlay.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBoxStopPlay
            // 
            this.textBoxStopPlay.Location = new System.Drawing.Point(211, 88);
            this.textBoxStopPlay.Name = "textBoxStopPlay";
            this.textBoxStopPlay.Size = new System.Drawing.Size(143, 21);
            this.textBoxStopPlay.TabIndex = 15;
            this.textBoxStopPlay.Text = "Ctrl+F12";
            this.textBoxStopPlay.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // notifyIcon_0
            // 
            this.notifyIcon_0.ContextMenuStrip = this.contextMenuStrip1;
            this.notifyIcon_0.Icon = ((System.Drawing.Icon)(resources.GetObject("notifyIcon_0.Icon")));
            this.notifyIcon_0.Text = "Dragon鼠标键盘模拟器";
            this.notifyIcon_0.Visible = true;
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.打开主窗体ToolStripMenuItem,
            this.访问官网ToolStripMenuItem,
            this.退出ToolStripMenuItem});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(137, 70);
            // 
            // 打开主窗体ToolStripMenuItem
            // 
            this.打开主窗体ToolStripMenuItem.Name = "打开主窗体ToolStripMenuItem";
            this.打开主窗体ToolStripMenuItem.Size = new System.Drawing.Size(136, 22);
            this.打开主窗体ToolStripMenuItem.Text = "打开主窗体";
            // 
            // 访问官网ToolStripMenuItem
            // 
            this.访问官网ToolStripMenuItem.Name = "访问官网ToolStripMenuItem";
            this.访问官网ToolStripMenuItem.Size = new System.Drawing.Size(136, 22);
            this.访问官网ToolStripMenuItem.Text = "访问官网";
            // 
            // 退出ToolStripMenuItem
            // 
            this.退出ToolStripMenuItem.Name = "退出ToolStripMenuItem";
            this.退出ToolStripMenuItem.Size = new System.Drawing.Size(136, 22);
            this.退出ToolStripMenuItem.Text = "退出";
            // 
            // buttonSetShotcut
            // 
            this.buttonSetShotcut.Location = new System.Drawing.Point(373, 115);
            this.buttonSetShotcut.Name = "buttonSetShotcut";
            this.buttonSetShotcut.Size = new System.Drawing.Size(115, 23);
            this.buttonSetShotcut.TabIndex = 18;
            this.buttonSetShotcut.Text = "保存快捷键设置";
            this.buttonSetShotcut.UseVisualStyleBackColor = true;
            // 
            // MacroForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.ClientSize = new System.Drawing.Size(503, 159);
            this.Controls.Add(this.textBoxStopPlay);
            this.Controls.Add(this.buttonSetShotcut);
            this.Controls.Add(this.textBoxPlay);
            this.Controls.Add(this.textBoxStop);
            this.Controls.Add(this.textBoxStart);
            this.Controls.Add(this.openbutton);
            this.Controls.Add(this.savebutton);
            this.Controls.Add(this.StopBackMacroButton);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.spannumericUpDown);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.loopnumericUpDown);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.playBackMacroButton);
            this.Controls.Add(this.recordStopButton);
            this.Controls.Add(this.recordStartButton);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "MacroForm";
            this.Text = "键盘鼠标模拟器";
            ((System.ComponentModel.ISupportInitialize)(this.loopnumericUpDown)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.spannumericUpDown)).EndInit();
            this.contextMenuStrip1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private void GlobalKeyboardHookKeyDown(object sender, KeyEventArgs e)
        {
            _macroEventsList.Add(new MacroEvent(MacroEventType.KeyDown, e, Environment.TickCount - _tickCount));
            _tickCount = Environment.TickCount;
        }

        private void GlobalKeyboardHookKeyUp(object sender, KeyEventArgs e)
        {
            _macroEventsList.Add(new MacroEvent(MacroEventType.KeyUp, e, Environment.TickCount - _tickCount));
            _tickCount = Environment.TickCount;
        }

        private void ShortcutKeyboardHookKeyDown(object sender, KeyEventArgs e)
        {
            if (_toolStatus != 0)
            {
                switch (_toolStatus)
                {
                    case 1:
                        textBoxStart.Text = e.KeyData.ToString();
                        _keysStartRec = (Keys)Enum.Parse(typeof(Keys), textBoxStart.Text, true);
                        break;
                    case 2:
                        textBoxStop.Text = e.KeyData.ToString();
                        _keysStopRec = (Keys)Enum.Parse(typeof(Keys), textBoxStop.Text, true);
                        break;
                    case 3:
                        textBoxPlay.Text = e.KeyData.ToString();
                        _keysStartPlay = (Keys)Enum.Parse(typeof(Keys), textBoxPlay.Text, true);
                        break;
                    case 4:
                        textBoxStopPlay.Text = e.KeyData.ToString();
                        _keysStopPlay = (Keys)Enum.Parse(typeof(Keys), textBoxStopPlay.Text, true);
                        break;
                }
            }
            else if (e.KeyData == _keysStartRec)
            {
                _macroEventsList.Clear();
                _tickCount = Environment.TickCount;
                _isPlayRecScript = false;
                _globalKeyboardHook.Start();
                _globalMouseHook.Start();
            }
            else if (e.KeyData == _keysStopRec)
            {
                _isPlayRecScript = false;
                _globalKeyboardHook.Stop();
                _globalMouseHook.Stop();
            }
            else if (e.KeyData == _keysStartPlay)
            {
                _isPlayRecScript = true;
                if (!BackgroundWorker.IsBusy)
                {
                    BackgroundWorker.RunWorkerAsync();
                }
            }
            else if (e.KeyData == _keysStopPlay)
            {
                _isPlayRecScript = false;
                if (BackgroundWorker.IsBusy)
                {
                    BackgroundWorker.CancelAsync();
                }
            }
        }

        private void ShortcutKeyboardHookKeyUp(object sender, KeyEventArgs e)
        {
            if (_toolStatus == 0)
            {
                if (e.KeyData == _keysStartRec)
                {
                    _macroEventsList.Clear();
                    _tickCount = Environment.TickCount;
                    _isPlayRecScript = false;
                    _globalKeyboardHook.Start();
                    _globalMouseHook.Start();
                }
                else if (e.KeyData == _keysStopRec)
                {
                    _isPlayRecScript = false;
                    _globalKeyboardHook.Stop();
                    _globalMouseHook.Stop();
                }
                else if (e.KeyData == _keysStartPlay)
                {
                    _isPlayRecScript = true;
                    if (!BackgroundWorker.IsBusy)
                    {
                        BackgroundWorker.RunWorkerAsync();
                    }
                }
                else if (e.KeyData == _keysStartRec)
                {
                    _isPlayRecScript = false;
                    if (BackgroundWorker.IsBusy)
                    {
                        BackgroundWorker.CancelAsync();
                    }
                }
            }
        }

        private void MacroForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            e.Cancel = true;
            Hide();//窗口最小化
            ShowInTaskbar = false;
            notifyIcon_0.Visible = true;
        }

        private void background_worker(object sender, DoWorkEventArgs e)
        {
            while (loopnumericUpDown.Value >= 0)
            {
                foreach (MacroEvent current in _macroEventsList)
                {
                    Thread.Sleep(current.TimeSinceLastEvent);
                    if (((BackgroundWorker)sender).CancellationPending)
                    {
                        e.Cancel = true;
                        break;
                    }
                    switch (current.MacroEventType)
                    {
                        case MacroEventType.MouseMove:
                            {
                                MouseEventArgs mouseEventArgs = (MouseEventArgs)current.EventArgs;
                                MouseSimulator.X = mouseEventArgs.X;
                                MouseSimulator.Y = mouseEventArgs.Y;
                                break;
                            }
                        case MacroEventType.MouseDown:
                            {
                                MouseEventArgs mouseEventArgs = (MouseEventArgs)current.EventArgs;
                                MouseSimulator.MouseDown(mouseEventArgs.Button);
                                break;
                            }
                        case MacroEventType.MouseUp:
                            {
                                MouseEventArgs mouseEventArgs = (MouseEventArgs)current.EventArgs;
                                MouseSimulator.MouseUp(mouseEventArgs.Button);
                                break;
                            }
                        case MacroEventType.MouseWheel:
                            {
                                MouseEventArgs mouseEventArgs = (MouseEventArgs)current.EventArgs;
                                MouseSimulator.MouseWheel(mouseEventArgs.Delta);
                                break;
                            }
                        case MacroEventType.KeyDown:
                            {
                                KeyEventArgs keyEventArgs = (KeyEventArgs)current.EventArgs;
                                KeyboardSimulator.KeyDown(keyEventArgs.KeyData);
                                break;
                            }
                        case MacroEventType.KeyUp:
                            {
                                KeyEventArgs keyEventArgs = (KeyEventArgs)current.EventArgs;
                                KeyboardSimulator.KeyUp(keyEventArgs.KeyData);
                                break;
                            }
                    }
                }
                Thread.Sleep(new TimeSpan(0, 0, (int)spannumericUpDown.Value));
                if (loopnumericUpDown.Value == 0)
                {
                    break;
                }
                loopnumericUpDown.Value--;
            }
        }

        private void GlobalMouseHookGlobalMouseDown(object sender, MouseEventArgs e)
        {
            _macroEventsList.Add(new MacroEvent(MacroEventType.MouseDown, e, Environment.TickCount - _tickCount));
            _tickCount = Environment.TickCount;
        }

        private void GlobalMouseHookGlobalMouseMove(object sender, MouseEventArgs e)
        {
            _macroEventsList.Add(new MacroEvent(MacroEventType.MouseMove, e, Environment.TickCount - _tickCount));
            _tickCount = Environment.TickCount;
        }
        private void GlobalMouseHookGlobalMouseUp(object sender, MouseEventArgs e)
        {
            _macroEventsList.Add(new MacroEvent(MacroEventType.MouseUp, e, Environment.TickCount - _tickCount));
            _tickCount = Environment.TickCount;
        }

        private void GlobalMouseHookGlobalMouseWheel(object sender, MouseEventArgs e)
        {
            _macroEventsList.Add(new MacroEvent(MacroEventType.MouseWheel, e, Environment.TickCount - _tickCount));
            _tickCount = Environment.TickCount;
        }
        private void NotifyIcon0MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                base.Show();
                ShowInTaskbar = true;
                notifyIcon_0.Visible = false;
            }
        }

        private void OpenbuttonClick(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog
                {
                    InitialDirectory = Directory.GetCurrentDirectory(),
                    Filter = "脚本文件(*.mic)|*.mic|所有文件|*.*",
                    RestoreDirectory = true,
                    FilterIndex = 1
                };
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string fileName = openFileDialog.FileName;
                    _macroEventsList.Clear();
                    _tickCount = Environment.TickCount;
                    XmlDocument xmlDocument = new XmlDocument();
                    xmlDocument.Load(fileName);
                    XmlNode xmlNode = xmlDocument.SelectSingleNode("MacroEvents");
                    if (xmlNode?.Attributes != null)
                    {
                        spannumericUpDown.Value = decimal.Parse(xmlNode.Attributes["TimeSpan"].Value);
                        spannumericUpDown.Value = decimal.Parse(xmlNode.Attributes["Loop"].Value);
                    }
                    var selectSingleNode = xmlDocument.SelectSingleNode("MacroEvents");
                    if (selectSingleNode != null)
                    {
                        XmlNodeList childNodes = selectSingleNode.ChildNodes;
                        foreach (XmlNode xmlNode2 in childNodes)
                        {
                            if (xmlNode2.Attributes != null)
                            {
                                string value = xmlNode2.Attributes["MET"].Value;
                                switch (value)
                                {
                                    case "MM":
                                        {
                                            int x = int.Parse(xmlNode2.Attributes["X"].Value);
                                            int y = int.Parse(xmlNode2.Attributes["Y"].Value);
                                            int timeSinceLastEvent = int.Parse(xmlNode2.Attributes["T"].Value);
                                            MouseEventArgs eventArgs = new MouseEventArgs(MouseButtons.None, 0, x, y, 0);
                                            MacroEvent item = new MacroEvent(MacroEventType.MouseMove, eventArgs, timeSinceLastEvent);
                                            _macroEventsList.Add(item);
                                            break;
                                        }
                                    case "MD":
                                        {
                                            int x = int.Parse(xmlNode2.Attributes["X"].Value);
                                            int y = int.Parse(xmlNode2.Attributes["Y"].Value);
                                            int clicks = int.Parse(xmlNode2.Attributes["C"].Value);
                                            int timeSinceLastEvent = int.Parse(xmlNode2.Attributes["T"].Value);
                                            string value2 = xmlNode2.Attributes["B"].Value;
                                            MouseButtons button = (MouseButtons)Enum.Parse(typeof(MouseButtons), value2, true);
                                            MouseEventArgs eventArgs = new MouseEventArgs(button, clicks, x, y, 0);
                                            MacroEvent item = new MacroEvent(MacroEventType.MouseDown, eventArgs, timeSinceLastEvent);
                                            _macroEventsList.Add(item);
                                            break;
                                        }
                                    case "MU":
                                        {
                                            int x = int.Parse(xmlNode2.Attributes["X"].Value);
                                            int y = int.Parse(xmlNode2.Attributes["Y"].Value);
                                            int clicks = int.Parse(xmlNode2.Attributes["C"].Value);
                                            int timeSinceLastEvent = int.Parse(xmlNode2.Attributes["T"].Value);
                                            string value2 = xmlNode2.Attributes["B"].Value;
                                            MouseButtons button = (MouseButtons)Enum.Parse(typeof(MouseButtons), value2, true);
                                            MouseEventArgs eventArgs = new MouseEventArgs(button, clicks, x, y, 0);
                                            MacroEvent item = new MacroEvent(MacroEventType.MouseUp, eventArgs, timeSinceLastEvent);
                                            _macroEventsList.Add(item);
                                            break;
                                        }
                                    case "MW":
                                        {
                                            int x = int.Parse(xmlNode2.Attributes["X"].Value);
                                            int y = int.Parse(xmlNode2.Attributes["Y"].Value);
                                            int delta = int.Parse(xmlNode2.Attributes["D"].Value);
                                            int timeSinceLastEvent = int.Parse(xmlNode2.Attributes["T"].Value);
                                            string value2 = xmlNode2.Attributes["B"].Value;
                                            MouseButtons button = (MouseButtons)Enum.Parse(typeof(MouseButtons), value2, true);
                                            MouseEventArgs eventArgs = new MouseEventArgs(button, 0, x, y, delta);
                                            MacroEvent item = new MacroEvent(MacroEventType.MouseWheel, eventArgs, timeSinceLastEvent);
                                            _macroEventsList.Add(item);
                                            break;
                                        }
                                    case "KD":
                                        {
                                            int timeSinceLastEvent = int.Parse(xmlNode2.Attributes["T"].Value);
                                            string value3 = xmlNode2.Attributes["K"].Value;
                                            Keys keyData = (Keys)Enum.Parse(typeof(Keys), value3, true);
                                            KeyEventArgs eventArgs2 = new KeyEventArgs(keyData);
                                            MacroEvent item = new MacroEvent(MacroEventType.KeyDown, eventArgs2, timeSinceLastEvent);
                                            _macroEventsList.Add(item);
                                            break;
                                        }
                                    case "KU":
                                        {
                                            int timeSinceLastEvent = int.Parse(xmlNode2.Attributes["T"].Value);
                                            string value3 = xmlNode2.Attributes["K"].Value;
                                            Keys keyData = (Keys)Enum.Parse(typeof(Keys), value3, true);
                                            KeyEventArgs eventArgs2 = new KeyEventArgs(keyData);
                                            MacroEvent item = new MacroEvent(MacroEventType.KeyUp, eventArgs2, timeSinceLastEvent);
                                            _macroEventsList.Add(item);
                                            break;
                                        }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void PlayBackMacroButtonClick(object sender, EventArgs e)
        {
            _isPlayRecScript = true;
            if (!BackgroundWorker.IsBusy)
            {
                BackgroundWorker.RunWorkerAsync();
            }
        }

        private void RecordStartButtonClick(object sender, EventArgs e)
        {
            _macroEventsList.Clear();
            _tickCount = Environment.TickCount;
            _globalKeyboardHook.Start();
            _globalMouseHook.Start();
        }

        private void RecordStopButtonClick(object sender, EventArgs e)
        {
            _globalKeyboardHook.Stop();
            _globalMouseHook.Stop();
        }
        private void SavebuttonClick(object sender, EventArgs e)
        {
            try
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog
                {
                    InitialDirectory = Directory.GetCurrentDirectory(),
                    Filter = "脚本文件(*.mic)|*.mic|所有文件|*.*",
                    RestoreDirectory = true,
                    FilterIndex = 1
                };
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string fileName = saveFileDialog.FileName;
                    XmlDocument xmlDocument = new XmlDocument();
                    XmlElement xmlElement = xmlDocument.CreateElement("MacroEvents");
                    xmlElement.SetAttribute("TimeSpan", spannumericUpDown.Value.ToString());
                    xmlElement.SetAttribute("Loop", loopnumericUpDown.Value.ToString());
                    xmlDocument.AppendChild(xmlElement);
                    int num = 0;
                    foreach (MacroEvent current in _macroEventsList)
                    {
                        num++;
                        XmlElement xmlElement2 = xmlDocument.CreateElement("ME" + num);
                        switch (current.MacroEventType)
                        {
                            case MacroEventType.MouseMove:
                                {
                                    xmlElement2.SetAttribute("MET", "MM");
                                    MouseEventArgs mouseEventArgs = (MouseEventArgs)current.EventArgs;
                                    xmlElement2.SetAttribute("X", mouseEventArgs.X.ToString());
                                    xmlElement2.SetAttribute("Y", mouseEventArgs.Y.ToString());
                                    xmlElement2.SetAttribute("T", current.TimeSinceLastEvent.ToString());
                                    xmlElement.AppendChild(xmlElement2);
                                    break;
                                }
                            case MacroEventType.MouseDown:
                                {
                                    xmlElement2.SetAttribute("MET", "MD");
                                    MouseEventArgs mouseEventArgs = (MouseEventArgs)current.EventArgs;
                                    xmlElement2.SetAttribute("X", mouseEventArgs.X.ToString());
                                    xmlElement2.SetAttribute("Y", mouseEventArgs.Y.ToString());
                                    xmlElement2.SetAttribute("C", mouseEventArgs.Clicks.ToString());
                                    xmlElement2.SetAttribute("B", mouseEventArgs.Button.ToString());
                                    xmlElement2.SetAttribute("T", current.TimeSinceLastEvent.ToString());
                                    xmlElement.AppendChild(xmlElement2);
                                    break;
                                }
                            case MacroEventType.MouseUp:
                                {
                                    xmlElement2.SetAttribute("MET", "MU");
                                    MouseEventArgs mouseEventArgs = (MouseEventArgs)current.EventArgs;
                                    xmlElement2.SetAttribute("C", mouseEventArgs.Clicks.ToString());
                                    xmlElement2.SetAttribute("X", mouseEventArgs.X.ToString());
                                    xmlElement2.SetAttribute("Y", mouseEventArgs.Y.ToString());
                                    xmlElement2.SetAttribute("B", mouseEventArgs.Button.ToString());
                                    xmlElement2.SetAttribute("T", current.TimeSinceLastEvent.ToString());
                                    xmlElement.AppendChild(xmlElement2);
                                    break;
                                }
                            case MacroEventType.MouseWheel:
                                {
                                    xmlElement2.SetAttribute("MET", "MW");
                                    MouseEventArgs mouseEventArgs = (MouseEventArgs)current.EventArgs;
                                    xmlElement2.SetAttribute("X", mouseEventArgs.X.ToString());
                                    xmlElement2.SetAttribute("Y", mouseEventArgs.Y.ToString());
                                    xmlElement2.SetAttribute("B", mouseEventArgs.Button.ToString());
                                    xmlElement2.SetAttribute("D", mouseEventArgs.Delta.ToString());
                                    xmlElement2.SetAttribute("T", current.TimeSinceLastEvent.ToString());
                                    xmlElement.AppendChild(xmlElement2);
                                    break;
                                }
                            case MacroEventType.KeyDown:
                                {
                                    xmlElement2.SetAttribute("MET", "KD");
                                    KeyEventArgs keyEventArgs = (KeyEventArgs)current.EventArgs;
                                    xmlElement2.SetAttribute("K", keyEventArgs.KeyData.ToString());
                                    xmlElement2.SetAttribute("T", current.TimeSinceLastEvent.ToString());
                                    xmlElement.AppendChild(xmlElement2);
                                    break;
                                }
                            case MacroEventType.KeyUp:
                                {
                                    xmlElement2.SetAttribute("MET", "KU");
                                    KeyEventArgs keyEventArgs = (KeyEventArgs)current.EventArgs;
                                    xmlElement2.SetAttribute("K", keyEventArgs.KeyData.ToString());
                                    xmlElement2.SetAttribute("T", current.TimeSinceLastEvent.ToString());
                                    xmlElement.AppendChild(xmlElement2);
                                    break;
                                }
                        }
                    }
                    xmlDocument.Save(fileName);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void StopBackMacroButton_Click(object sender, EventArgs e)
        {
            _isPlayRecScript = false;
            _globalKeyboardHook.Stop();
            _globalMouseHook.Stop();
        }
        private void TextBoxPlayMouseHover(object sender, EventArgs e)
        {
            _toolStatus = 3;
        }

        private void TextBoxPlayMouseLeave(object sender, EventArgs e)
        {
            _toolStatus = 0;
        }

        private void TextBoxStartMouseHover(object sender, EventArgs e)
        {
            _toolStatus = 1;
        }

        private void TextBoxStartMouseLeave(object sender, EventArgs e)
        {
            _toolStatus = 0;
        }

        private void TextBoxStopMouseHover(object sender, EventArgs e)
        {
            _toolStatus = 2;
        }
        private void TextBoxStopMouseLeave(object sender, EventArgs e)
        {
            _toolStatus = 0;
        }

        private void TextBoxStopPlayMouseHover(object sender, EventArgs e)
        {
            _toolStatus = 4;
        }
        private void TextBoxStopPlayMouseLeave(object sender, EventArgs e)
        {
            _toolStatus = 0;
        }

        private void ButtonSetShotcutOnClick(object sender, EventArgs e)
        {
            _toolStatus = 0;
            MessageBox.Show("快捷键设置已经保存！");
        }

        private void 打开主窗体ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            base.Show();
            ShowInTaskbar = true;
            notifyIcon_0.Visible = false;
        }

        private void 访问官网ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RegistryKey registryKey = Registry.ClassesRoot.OpenSubKey("http\\shell\\open\\command\\");
            if (registryKey != null)
            {
                string text = registryKey.GetValue("").ToString();
                text = text.Substring(1, text.Length - 10);
                Process.Start(text, "http://www.cnblogs.com/Chary/");
            }
        }

        private void 退出ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            _isPlayRecScript = false;
            _shortcutKeyboardHook.Stop();
            _globalKeyboardHook.Stop();
            _globalMouseHook.Stop();
            if (BackgroundWorker.IsBusy)
            {
                BackgroundWorker.CancelAsync();
            }
            Dispose(true);
            Application.ExitThread();
        }

        #endregion Private Methods
    }
}