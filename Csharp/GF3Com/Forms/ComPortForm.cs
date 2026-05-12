namespace GF3Com.Forms;

/// <summary>
/// COM Port 監控介面 (frmComPort.frm)
/// 顯示各 PLC port 的接收字串、長度與狀態。
/// </summary>
internal partial class ComPortForm : Form
{
    // 對應 VB6 的 lblComPortString(n), lblComPortLen(n), lblComPortStatus(n)
    // n 範圍：21–25 (COM4+17 ~ COM8+17)
    private readonly Label[] _lblString = new Label[26];
    private readonly Label[] _lblLen    = new Label[26];
    private readonly Label[] _lblStatus = new Label[26];

    // AGV 輸出字串 (index 9)
    private readonly Label[] _lblOutString = new Label[10];
    private readonly Label[] _lblOutLen    = new Label[10];

    // AGV 狀態 (index 1–2)
    private readonly Label[] _lblAgvStatus = new Label[3];

    public ComPortForm()
    {
        InitializeComponent();
    }

    // ── 公開更新方法（供 PlcCommunicator 呼叫）────────────────────

    /// <summary>
    /// 更新 PLC port 顯示 (對應 frmComPort.lblComPortXxx(index))
    /// index 範圍：21–25
    /// </summary>
    public void UpdatePortDisplay(int index, string len, string data,
                                   string status, Color statusColor)
    {
        if (!IsHandleCreated) return;
        this.BeginInvoke(() =>
        {
            if (_lblLen[index]    is not null) _lblLen[index].Text    = len;
            if (_lblString[index] is not null) _lblString[index].Text = data;
            if (_lblStatus[index] is not null)
            {
                _lblStatus[index].Text      = status;
                _lblStatus[index].ForeColor = statusColor;
            }
        });
    }

    /// <summary>
    /// 更新 AGV 輸出顯示 (index 9)
    /// </summary>
    public void UpdateAgvOutput(string len, string data)
    {
        if (!IsHandleCreated) return;
        this.BeginInvoke(() =>
        {
            if (_lblOutLen[9]    is not null) _lblOutLen[9].Text    = len;
            if (_lblOutString[9] is not null) _lblOutString[9].Text = data;
        });
    }

    /// <summary>
    /// 更新 AGV 狀態 (index 1–2)
    /// </summary>
    public void UpdateAgvStatus(int index, string status)
    {
        if (!IsHandleCreated || index < 1 || index > 2) return;
        this.BeginInvoke(() =>
        {
            if (_lblAgvStatus[index] is not null)
                _lblAgvStatus[index].Text = status;
        });
    }

    // ── 事件處理 ─────────────────────────────────────────────────────

    private void CmdEnd_Click(object? sender, EventArgs e) => Close();
}
