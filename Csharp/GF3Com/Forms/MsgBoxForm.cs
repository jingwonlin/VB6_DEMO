namespace GF3Com.Forms;

/// <summary>
/// 自訂訊息框 (frmMsgBox.frm)
/// 支援自動關閉 (autoCloseMs)，對應 VB6 xMsgBox 的 lTime 參數。
/// </summary>
internal partial class MsgBoxForm : Form
{
    private System.Windows.Forms.Timer? _autoCloseTimer;
    private DialogResult _result = DialogResult.OK;

    public MsgBoxForm()
    {
        InitializeComponent();
    }

    /// <summary>
    /// 顯示訊息並回傳使用者選擇結果
    /// </summary>
    public static DialogResult Show(string message,
                                     MessageBoxButtons buttons = MessageBoxButtons.OK,
                                     string title = "",
                                     int autoCloseMs = 0)
    {
        using var frm = new MsgBoxForm();
        frm.Setup(message, buttons, title, autoCloseMs);
        frm.ShowDialog();
        return frm._result;
    }

    private void Setup(string message, MessageBoxButtons buttons,
                        string title, int autoCloseMs)
    {
        lblMessage.Text = message;
        Text = string.IsNullOrEmpty(title) ? Application.ProductName : title;

        // 依 buttons 設定按鈕可見性
        btnOk.Visible     = buttons is MessageBoxButtons.OK or MessageBoxButtons.OKCancel;
        btnCancel.Visible = buttons is MessageBoxButtons.OKCancel or MessageBoxButtons.YesNoCancel;
        btnYes.Visible    = buttons is MessageBoxButtons.YesNo or MessageBoxButtons.YesNoCancel;
        btnNo.Visible     = buttons is MessageBoxButtons.YesNo or MessageBoxButtons.YesNoCancel;

        if (autoCloseMs > 0)
        {
            _autoCloseTimer = new System.Windows.Forms.Timer { Interval = autoCloseMs };
            _autoCloseTimer.Tick += (_, _) =>
            {
                _result = DialogResult.OK;
                Close();
            };
            _autoCloseTimer.Start();
        }
    }

    private void BtnOk_Click(object? sender, EventArgs e)
    {
        _result = DialogResult.OK;
        Close();
    }

    private void BtnCancel_Click(object? sender, EventArgs e)
    {
        _result = DialogResult.Cancel;
        Close();
    }

    private void BtnYes_Click(object? sender, EventArgs e)
    {
        _result = DialogResult.Yes;
        Close();
    }

    private void BtnNo_Click(object? sender, EventArgs e)
    {
        _result = DialogResult.No;
        Close();
    }

    protected override void OnFormClosed(FormClosedEventArgs e)
    {
        _autoCloseTimer?.Stop();
        _autoCloseTimer?.Dispose();
        base.OnFormClosed(e);
    }
}
