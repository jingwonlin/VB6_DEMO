namespace GF3Com.Forms;

/// <summary>
/// 應用程式啟動訊息表單 (frmApStartMsg.frm)
/// </summary>
internal partial class StartupMsgForm : Form
{
    public StartupMsgForm()
    {
        InitializeComponent();
    }

    public string Message
    {
        get => lblMessage.Text;
        set => lblMessage.Text = value;
    }
}
