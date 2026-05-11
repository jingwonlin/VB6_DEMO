namespace GF3Com.Forms;

partial class StartupMsgForm
{
    private System.ComponentModel.IContainer components = null;
    internal Label lblMessage = null!;

    protected override void Dispose(bool disposing)
    {
        if (disposing && components != null)
            components.Dispose();
        base.Dispose(disposing);
    }

    private void InitializeComponent()
    {
        lblMessage = new Label();
        SuspendLayout();

        lblMessage.AutoSize  = false;
        lblMessage.Dock      = DockStyle.Fill;
        lblMessage.TextAlign = ContentAlignment.MiddleCenter;
        lblMessage.Font      = new Font("新細明體", 14f);
        lblMessage.Text      = "應用程式啟動中，請稍候...";

        ClientSize = new Size(480, 120);
        Controls.Add(lblMessage);
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox     = false;
        MinimizeBox     = false;
        StartPosition   = FormStartPosition.CenterScreen;
        Text            = "啟動";
        ResumeLayout(false);
    }
}
