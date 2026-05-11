using System.IO.Ports;
using GF3Com.Communication;
using GF3Com.Database;
using GF3Com.Utilities;

namespace GF3Com.Forms;

/// <summary>
/// 主 MDI 視窗 (MDImainForm.frm)
/// </summary>
internal partial class MainForm : Form
{
    // ── 通訊物件 ────────────────────────────────────────────────────
    private readonly SerialPortWrapper[] _commPorts = new SerialPortWrapper[10]; // index 4–9
    private PlcCommunicator? _plc;
    private AgvCommunicator? _agv;
    private readonly DatabaseContext _db = new();
    private ComPortForm? _comPortForm;

    // ── Timers ───────────────────────────────────────────────────────
    private System.Windows.Forms.Timer _timerNew  = null!;
    private System.Windows.Forms.Timer _timerAgv  = null!;

    // ── 局部狀態 ────────────────────────────────────────────────────
    private int _portIndex = 4;   // iPortIndex in VB6
    private int _bccErrCnt;
    private readonly bool[] _lineNo = new bool[25];

    public MainForm()
    {
        InitializeComponent();
    }

    // ── 表單載入 ─────────────────────────────────────────────────────

    private void MainForm_Load(object sender, EventArgs e)
    {
        try
        {
            SetStatusBar(1, Name);

            using var splash = new StartupMsgForm();
            splash.Message = $"{Text} 啟動中，請稍候...";
            splash.Show();
            Application.DoEvents();

            // 單一實例檢查
            if (System.Diagnostics.Process.GetProcessesByName(
                    System.IO.Path.GetFileNameWithoutExtension(
                        Application.ExecutablePath)).Length > 1)
            {
                AppUtility.ShowMessage($"[{Application.ProductName}] 已在執行中！");
                Application.Exit();
                return;
            }

            // 確認 IniFile 目錄存在
            string iniDir = Path.Combine(Application.StartupPath, "IniFile");
            string iniFile= Path.Combine(iniDir, "GF3Com.ini");
            if (!Directory.Exists(iniDir))
            {
                Directory.CreateDirectory(iniDir);
                AppUtility.ShowMessage(
                    "IniFile 子目錄中找不到 GF3Com.ini！",
                    MessageBoxButtons.OK, MessageBoxIcon.Error, "Error Message");
                Application.Exit();
                return;
            }

            string logDir = Path.Combine(Application.StartupPath, "LogMessage");
            Directory.CreateDirectory(logDir);

            SetStatusBar(3, "資料庫連線中");
            DataBaseConnect();
            Globals.InitialParameter();
            InitParameter1();

            // 開啟 PLC COM ports (COM4–COM8)
            for (int n = 4; n <= 8; n++)
            {
                _commPorts[n] = new SerialPortWrapper(n);
                try
                {
                    _commPorts[n].Open(9600, Parity.Even, 8, StopBits.One,
                        n switch { 4 => 28, 5 => 16, 6 => 12, 7 => 16, 8 => 50, _ => 1 });
                    Globals.ComOpenOk[n] = 1;
                }
                catch
                {
                    Globals.ComOpenOk[n] = -1;
                }
            }

            // AGV port (COM9) — 依 ini 設定決定是否啟用
            _commPorts[9] = new SerialPortWrapper(9);
            _agv = new AgvCommunicator(_commPorts[9]);

            // 建立 PlcCommunicator
            _plc = new PlcCommunicator(
                _commPorts,
                _db,
                (idx, len, str, stat, col) =>
                {
                    if (_comPortForm?.Visible == true)
                        _comPortForm.UpdatePortDisplay(idx, len, str, stat, col);
                });

            // Timers
            _timerNew.Interval = 600;
            _timerNew.Enabled  = true;

            _timerAgv.Interval = 2000;
            _timerAgv.Enabled  = false;

            splash.Close();
        }
        catch (Exception ex)
        {
            AppUtility.ShowMessage(
                $"[MDImainForm_Load] 錯誤：{ex.Message}",
                MessageBoxButtons.OK, MessageBoxIcon.Error, "Error", 5000);
        }
    }

    // ── 資料庫連線 ───────────────────────────────────────────────────

    private void DataBaseConnect()
    {
        Globals.ComputerName = AppUtility.GetComputerName();

        if (string.Equals(Globals.ComputerName[..Math.Min(8, Globals.ComputerName.Length)],
                          "TPNPCENG", StringComparison.OrdinalIgnoreCase))
            Globals.ComputerName = "GF2";

        if (Globals.ComputerName.StartsWith("GF2", StringComparison.OrdinalIgnoreCase)
            || Globals.ComputerName.StartsWith("SKNPCEMD195", StringComparison.OrdinalIgnoreCase))
        {
            Globals.SqlConnectionString =
                "Data Source=10.115.80.195;Initial Catalog=sqldb;User ID=sa;Password=nanya;TrustServerCertificate=True";
            Globals.ProjectName = "SK01";
            Globals.PlantName   = "南亞二";
        }
        else if (Globals.ComputerName.StartsWith("SKGF3", StringComparison.OrdinalIgnoreCase))
        {
            Globals.SqlConnectionString =
                "Data Source=192.115.64.77;Initial Catalog=sqldb;User ID=sa;Password=nanya;TrustServerCertificate=True";
            Globals.ProjectName = "SK02";
            Globals.PlantName   = "南亞三";
        }

        if (Globals.ComputerName.StartsWith("TPNPCENG", StringComparison.OrdinalIgnoreCase))
        {
            Globals.SqlConnectionStringLocal =
                "Data Source=.;Initial Catalog=sqldb;User ID=sa;TrustServerCertificate=True";
            _db.Open(Globals.SqlConnectionStringLocal);
        }
        else
        {
            _db.Open(Globals.SqlConnectionString);
        }
    }

    private void InitParameter1()
    {
        Globals.CranMoveFlag   = true;
        Globals.BackupDir      = @"c:\";
        Globals.Refresh1Screen = 0;
        Globals.DbType         = "SQL";
        Globals.Plc1Num        = 7;
        Globals.StnoNum        = 5;
        Globals.Err2Cnt        = 0;
        Globals.LoginUser      = AppUtility.FixedLen("", 10);
        Globals.LoginPass      = AppUtility.FixedLen("", 5);
        Globals.LoginName      = AppUtility.FixedLen("", 8);
        Globals.Super          = "N";
    }

    // ── Timer 事件 ───────────────────────────────────────────────────

    private void TimerNew_Tick(object? sender, EventArgs e)
    {
        try
        {
            _timerNew.Enabled = false;
            _plc?.Poll();
        }
        catch (Exception ex)
        {
            AppUtility.ShowMessage(
                $"[TimerNew_Tick] 錯誤：{ex.Message}",
                MessageBoxButtons.OK, MessageBoxIcon.Error, "Error", 5000);
        }
        finally
        {
            _timerNew.Enabled = true;
        }
    }

    private void TimerAgv_Tick(object? sender, EventArgs e)
    {
        // AGV 週期性狀態查詢（保留擴充點）
    }

    // ── 選單 ─────────────────────────────────────────────────────────

    private void MenuComPort_Click(object? sender, EventArgs e)
    {
        _comPortForm ??= new ComPortForm();
        if (_comPortForm.IsDisposed)
            _comPortForm = new ComPortForm();
        _comPortForm.MdiParent = this;
        _comPortForm.Show();
        _comPortForm.BringToFront();
    }

    private void MenuExit_Click(object? sender, EventArgs e)
    {
        Close();
        Application.Exit();
    }

    // ── 視窗關閉 ─────────────────────────────────────────────────────

    private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
    {
        // 只允許從選單正常離開 (UnloadMode != 1 → e.CloseReason != UserClosing in VB6)
        if (e.CloseReason != CloseReason.UserClosing &&
            e.CloseReason != CloseReason.ApplicationExitCall)
        {
            AppUtility.ShowMessage(
                "無法在此模式下結束作業系統",
                MessageBoxButtons.OK, MessageBoxIcon.Warning, "操作提示", 5000);
            e.Cancel = true;
        }
    }

    private void MainForm_FormClosed(object sender, FormClosedEventArgs e)
    {
        _timerNew.Enabled = false;
        _timerAgv.Enabled = false;
        for (int n = 4; n <= 9; n++)
            _commPorts[n]?.Dispose();
        _db.Dispose();
    }

    // ── 工具方法 ─────────────────────────────────────────────────────

    private void SetStatusBar(int panelIndex, string text)
    {
        AppUtility.SetStatusBar(ctlStatusBar, panelIndex, text);
    }
}
