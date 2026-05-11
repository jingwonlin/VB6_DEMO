namespace GF3Com.Forms;

partial class MainForm
{
    private System.ComponentModel.IContainer components = null;

    // 狀態列
    internal StatusStrip ctlStatusBar = null!;
    private ToolStripStatusLabel _panelOp   = null!;
    private ToolStripStatusLabel _panelInfo = null!;
    private ToolStripStatusLabel _panelSt   = null!;
    private ToolStripStatusLabel _panelConn = null!;
    private ToolStripStatusLabel _panelDate = null!;
    private ToolStripStatusLabel _panelTime = null!;

    // 底部資訊列
    private Panel _bottomPanel = null!;
    private Label _lblOperMode = null!;

    // 選單
    private MenuStrip _menuStrip    = null!;
    private ToolStripMenuItem _mnuFunction = null!;
    private ToolStripMenuItem _mnuComPort  = null!;
    private ToolStripMenuItem _mnuSep      = null!;
    private ToolStripMenuItem _mnuExit     = null!;

    protected override void Dispose(bool disposing)
    {
        if (disposing && components != null)
            components.Dispose();
        base.Dispose(disposing);
    }

    private void InitializeComponent()
    {
        // ── 選單 ────────────────────────────────────────────────────
        _menuStrip    = new MenuStrip();
        _mnuFunction  = new ToolStripMenuItem("功能選項(&1)");
        _mnuComPort   = new ToolStripMenuItem("串聯動態");
        _mnuSep       = new ToolStripSeparator() as ToolStripMenuItem ?? new ToolStripMenuItem("-");
        _mnuExit      = new ToolStripMenuItem("結束作業");

        _mnuComPort.Click += MenuComPort_Click;
        _mnuExit.Click    += MenuExit_Click;

        _mnuFunction.DropDownItems.Add(_mnuComPort);
        _mnuFunction.DropDownItems.Add(new ToolStripSeparator());
        _mnuFunction.DropDownItems.Add(_mnuExit);
        _menuStrip.Items.Add(_mnuFunction);

        // ── 狀態列 ──────────────────────────────────────────────────
        ctlStatusBar = new StatusStrip();
        _panelOp     = new ToolStripStatusLabel { Width = 320,  BorderSides = ToolStripStatusLabelBorderSides.Right, AutoSize = false };
        _panelInfo   = new ToolStripStatusLabel { Width = 160,  BorderSides = ToolStripStatusLabelBorderSides.Right, AutoSize = false, Text = "操作確認" };
        _panelSt     = new ToolStripStatusLabel { Width = 440,  BorderSides = ToolStripStatusLabelBorderSides.Right, AutoSize = false };
        _panelConn   = new ToolStripStatusLabel { Width = 220,  BorderSides = ToolStripStatusLabelBorderSides.Right, AutoSize = false, Text = "None" };
        _panelDate   = new ToolStripStatusLabel { Spring = false, AutoSize = false, Width = 180 };
        _panelTime   = new ToolStripStatusLabel { Spring = false, AutoSize = false, Width = 185 };

        ctlStatusBar.Items.AddRange(new ToolStripItem[]
            { _panelOp, _panelInfo, _panelSt, _panelConn, _panelDate, _panelTime });
        ctlStatusBar.Font = new Font("新細明體", 12f);

        // ── 底部資訊面板 ──────────────────────────────────────────
        _bottomPanel = new Panel
        {
            Dock   = DockStyle.Bottom,
            Height = 36,
            BackColor = SystemColors.Control
        };
        _lblOperMode = new Label
        {
            Text      = "操作模式",
            Dock      = DockStyle.Left,
            Width     = 90,
            TextAlign = ContentAlignment.MiddleCenter,
            Font      = new Font("新細明體", 12f),
            BackColor = Color.Cyan,
            ForeColor = Color.FromArgb(0xFF, 0x80, 0x80)
        };
        _bottomPanel.Controls.Add(_lblOperMode);

        // ── Timers ───────────────────────────────────────────────
        _timerNew = new System.Windows.Forms.Timer { Interval = 600, Enabled = false };
        _timerNew.Tick += TimerNew_Tick;

        _timerAgv = new System.Windows.Forms.Timer { Interval = 5000, Enabled = false };
        _timerAgv.Tick += TimerAgv_Tick;

        // ── 主表單屬性 ───────────────────────────────────────────
        IsMdiContainer    = true;
        BackColor         = SystemColors.AppWorkspace;
        ClientSize        = new Size(990, 680);
        Text              = "南亞科技倉儲自動化通訊系統連線";
        WindowState       = FormWindowState.Maximized;
        MainMenuStrip     = _menuStrip;

        Controls.Add(_menuStrip);
        Controls.Add(ctlStatusBar);
        Controls.Add(_bottomPanel);

        Load         += MainForm_Load;
        FormClosing  += MainForm_FormClosing;
        FormClosed   += MainForm_FormClosed;

        ResumeLayout(false);
        PerformLayout();
    }
}
