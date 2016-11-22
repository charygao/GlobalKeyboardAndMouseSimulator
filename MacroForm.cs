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
            components = new Container();
            ComponentResourceManager resources = new ComponentResourceManager(typeof(MacroForm));
            recordStartButton = new Button();
            recordStopButton = new Button();
            playBackMacroButton = new Button();
            label1 = new Label();
            label2 = new Label();
            loopnumericUpDown = new NumericUpDown();
            label3 = new Label();
            label4 = new Label();
            spannumericUpDown = new NumericUpDown();
            label5 = new Label();
            label6 = new Label();
            StopBackMacroButton = new Button();
            savebutton = new Button();
            openbutton = new Button();
            textBoxStart = new TextBox();
            textBoxStop = new TextBox();
            textBoxPlay = new TextBox();
            textBoxStopPlay = new TextBox();
            notifyIcon_0 = new NotifyIcon(components);
            contextMenuStrip1 = new ContextMenuStrip(components);
            打开主窗体ToolStripMenuItem = new ToolStripMenuItem();
            访问官网ToolStripMenuItem = new ToolStripMenuItem();
            退出ToolStripMenuItem = new ToolStripMenuItem();
            buttonSetShotcut = new Button();
            ((ISupportInitialize)(loopnumericUpDown)).BeginInit();
            ((ISupportInitialize)(spannumericUpDown)).BeginInit();
            contextMenuStrip1.SuspendLayout();
            SuspendLayout();
            // 
            // recordStartButton
            // 
            recordStartButton.Location = new System.Drawing.Point(52, 15);
            recordStartButton.Name = "recordStartButton";
            recordStartButton.Size = new System.Drawing.Size(140, 21);
            recordStartButton.TabIndex = 0;
            recordStartButton.Text = "开始录制(Ctrl+F9)";
            recordStartButton.UseVisualStyleBackColor = true;
            // 
            // recordStopButton
            // 
            recordStopButton.Location = new System.Drawing.Point(211, 15);
            recordStopButton.Name = "recordStopButton";
            recordStopButton.Size = new System.Drawing.Size(143, 21);
            recordStopButton.TabIndex = 0;
            recordStopButton.Text = "停止录制(Ctrl+F10)";
            recordStopButton.UseVisualStyleBackColor = true;
            // 
            // playBackMacroButton
            // 
            playBackMacroButton.Location = new System.Drawing.Point(52, 65);
            playBackMacroButton.Name = "playBackMacroButton";
            playBackMacroButton.Size = new System.Drawing.Size(140, 21);
            playBackMacroButton.TabIndex = 1;
            playBackMacroButton.Text = "开始播放(Ctrl+F11)";
            playBackMacroButton.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new System.Drawing.Point(17, 19);
            label1.Name = "label1";
            label1.Size = new System.Drawing.Size(29, 12);
            label1.TabIndex = 2;
            label1.Text = "录制";
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new System.Drawing.Point(17, 69);
            label2.Name = "label2";
            label2.Size = new System.Drawing.Size(29, 12);
            label2.TabIndex = 2;
            label2.Text = "播放";
            // 
            // loopnumericUpDown
            // 
            loopnumericUpDown.Location = new System.Drawing.Point(429, 65);
            loopnumericUpDown.Maximum = new decimal(new int[] {
            -159383553,
            46653770,
            5421,
            0});
            loopnumericUpDown.Name = "loopnumericUpDown";
            loopnumericUpDown.Size = new System.Drawing.Size(40, 21);
            loopnumericUpDown.TabIndex = 3;
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Location = new System.Drawing.Point(371, 69);
            label3.Name = "label3";
            label3.Size = new System.Drawing.Size(53, 12);
            label3.TabIndex = 4;
            label3.Text = "重复次数";
            // 
            // label4
            // 
            label4.AutoSize = true;
            label4.Location = new System.Drawing.Point(371, 24);
            label4.Name = "label4";
            label4.Size = new System.Drawing.Size(53, 12);
            label4.TabIndex = 5;
            label4.Text = "播放间隔";
            // 
            // spannumericUpDown
            // 
            spannumericUpDown.Location = new System.Drawing.Point(429, 20);
            spannumericUpDown.Maximum = new decimal(new int[] {
            -1593835521,
            466537709,
            54210,
            0});
            spannumericUpDown.Name = "spannumericUpDown";
            spannumericUpDown.Size = new System.Drawing.Size(40, 21);
            spannumericUpDown.TabIndex = 6;
            // 
            // label5
            // 
            label5.AutoSize = true;
            label5.Location = new System.Drawing.Point(474, 24);
            label5.Name = "label5";
            label5.Size = new System.Drawing.Size(17, 12);
            label5.TabIndex = 7;
            label5.Text = "秒";
            // 
            // label6
            // 
            label6.AutoSize = true;
            label6.Location = new System.Drawing.Point(474, 69);
            label6.Name = "label6";
            label6.Size = new System.Drawing.Size(17, 12);
            label6.TabIndex = 8;
            label6.Text = "次";
            // 
            // StopBackMacroButton
            // 
            StopBackMacroButton.Location = new System.Drawing.Point(211, 65);
            StopBackMacroButton.Name = "StopBackMacroButton";
            StopBackMacroButton.Size = new System.Drawing.Size(143, 21);
            StopBackMacroButton.TabIndex = 9;
            StopBackMacroButton.Text = "停止播放(Ctrl+F12)";
            StopBackMacroButton.UseVisualStyleBackColor = true;
            // 
            // savebutton
            // 
            savebutton.Location = new System.Drawing.Point(52, 115);
            savebutton.Name = "savebutton";
            savebutton.Size = new System.Drawing.Size(140, 23);
            savebutton.TabIndex = 10;
            savebutton.Text = "保存脚本";
            savebutton.UseVisualStyleBackColor = true;
            // 
            // openbutton
            // 
            openbutton.Location = new System.Drawing.Point(211, 115);
            openbutton.Name = "openbutton";
            openbutton.Size = new System.Drawing.Size(143, 23);
            openbutton.TabIndex = 11;
            openbutton.Text = "打开脚本";
            openbutton.UseVisualStyleBackColor = true;
            // 
            // textBoxStart
            // 
            textBoxStart.Location = new System.Drawing.Point(52, 41);
            textBoxStart.Name = "textBoxStart";
            textBoxStart.Size = new System.Drawing.Size(140, 21);
            textBoxStart.TabIndex = 12;
            textBoxStart.Text = "Ctrl+F9";
            textBoxStart.TextAlign = HorizontalAlignment.Center;
            // 
            // textBoxStop
            // 
            textBoxStop.Location = new System.Drawing.Point(211, 42);
            textBoxStop.Name = "textBoxStop";
            textBoxStop.Size = new System.Drawing.Size(143, 21);
            textBoxStop.TabIndex = 13;
            textBoxStop.Text = "Ctrl+F10";
            textBoxStop.TextAlign = HorizontalAlignment.Center;
            // 
            // textBoxPlay
            // 
            textBoxPlay.Location = new System.Drawing.Point(52, 88);
            textBoxPlay.Name = "textBoxPlay";
            textBoxPlay.Size = new System.Drawing.Size(140, 21);
            textBoxPlay.TabIndex = 14;
            textBoxPlay.Text = "Ctrl+F11";
            textBoxPlay.TextAlign = HorizontalAlignment.Center;
            // 
            // textBoxStopPlay
            // 
            textBoxStopPlay.Location = new System.Drawing.Point(211, 88);
            textBoxStopPlay.Name = "textBoxStopPlay";
            textBoxStopPlay.Size = new System.Drawing.Size(143, 21);
            textBoxStopPlay.TabIndex = 15;
            textBoxStopPlay.Text = "Ctrl+F12";
            textBoxStopPlay.TextAlign = HorizontalAlignment.Center;
            // 
            // notifyIcon_0
            // 
            notifyIcon_0.ContextMenuStrip = contextMenuStrip1;
            notifyIcon_0.Icon = ((System.Drawing.Icon)(resources.GetObject("notifyIcon_0.Icon")));
            notifyIcon_0.Text = "Dragon鼠标键盘模拟器";
            notifyIcon_0.Visible = true;
            // 
            // contextMenuStrip1
            // 
            contextMenuStrip1.Items.AddRange(new ToolStripItem[] {
            打开主窗体ToolStripMenuItem,
            访问官网ToolStripMenuItem,
            退出ToolStripMenuItem});
            contextMenuStrip1.Name = "contextMenuStrip1";
            contextMenuStrip1.Size = new System.Drawing.Size(137, 70);
            // 
            // 打开主窗体ToolStripMenuItem
            // 
            打开主窗体ToolStripMenuItem.Name = "打开主窗体ToolStripMenuItem";
            打开主窗体ToolStripMenuItem.Size = new System.Drawing.Size(136, 22);
            打开主窗体ToolStripMenuItem.Text = "打开主窗体";
            // 
            // 访问官网ToolStripMenuItem
            // 
            访问官网ToolStripMenuItem.Name = "访问官网ToolStripMenuItem";
            访问官网ToolStripMenuItem.Size = new System.Drawing.Size(136, 22);
            访问官网ToolStripMenuItem.Text = "访问官网";
            // 
            // 退出ToolStripMenuItem
            // 
            退出ToolStripMenuItem.Name = "退出ToolStripMenuItem";
            退出ToolStripMenuItem.Size = new System.Drawing.Size(136, 22);
            退出ToolStripMenuItem.Text = "退出";
            // 
            // buttonSetShotcut
            // 
            buttonSetShotcut.Location = new System.Drawing.Point(373, 115);
            buttonSetShotcut.Name = "buttonSetShotcut";
            buttonSetShotcut.Size = new System.Drawing.Size(115, 23);
            buttonSetShotcut.TabIndex = 18;
            buttonSetShotcut.Text = "保存快捷键设置";
            buttonSetShotcut.UseVisualStyleBackColor = true;
            // 
            // MacroForm
            // 
            AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = System.Drawing.SystemColors.Control;
            ClientSize = new System.Drawing.Size(503, 159);
            Controls.Add(textBoxStopPlay);
            Controls.Add(buttonSetShotcut);
            Controls.Add(textBoxPlay);
            Controls.Add(textBoxStop);
            Controls.Add(textBoxStart);
            Controls.Add(openbutton);
            Controls.Add(savebutton);
            Controls.Add(StopBackMacroButton);
            Controls.Add(label6);
            Controls.Add(label5);
            Controls.Add(spannumericUpDown);
            Controls.Add(label4);
            Controls.Add(label3);
            Controls.Add(loopnumericUpDown);
            Controls.Add(label2);
            Controls.Add(label1);
            Controls.Add(playBackMacroButton);
            Controls.Add(recordStopButton);
            Controls.Add(recordStartButton);
            Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            MaximizeBox = false;
            Name = "MacroForm";
            Text = "键盘鼠标模拟器";
            FormClosing += MacroForm_FormClosing;
            ((ISupportInitialize)(loopnumericUpDown)).EndInit();
            ((ISupportInitialize)(spannumericUpDown)).EndInit();
            contextMenuStrip1.ResumeLayout(false);
            ResumeLayout(false);
            PerformLayout();

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