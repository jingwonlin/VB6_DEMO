namespace GF3Com.Forms;

partial class MsgBoxForm
{
    private System.ComponentModel.IContainer components = null;
    private Label   lblMessage = null!;
    private Button  btnOk      = null!;
    private Button  btnCancel  = null!;
    private Button  btnYes     = null!;
    private Button  btnNo      = null!;

    protected override void Dispose(bool disposing)
    {
        if (disposing && components != null)
            components.Dispose();
        base.Dispose(disposing);
    }

    private void InitializeComponent()
    {
        lblMessage = new Label();
        btnOk      = new Button();
        btnCancel  = new Button();
        btnYes     = new Button();
        btnNo      = new Button();
        SuspendLayout();

        // lblMessage
        lblMessage.Location  = new Point(12, 12);
        lblMessage.Size      = new Size(460, 80);
        lblMessage.AutoSize  = false;
        lblMessage.TextAlign = ContentAlignment.MiddleLeft;
        lblMessage.Font      = new Font("新細明體", 12f);

        // btnOk
        btnOk.Text     = "確定";
        btnOk.Location = new Point(100, 105);
        btnOk.Size     = new Size(80, 30);
        btnOk.Click   += BtnOk_Click;

        // btnCancel
        btnCancel.Text     = "取消";
        btnCancel.Location = new Point(200, 105);
        btnCancel.Size     = new Size(80, 30);
        btnCancel.Click   += BtnCancel_Click;

        // btnYes
        btnYes.Text     = "是";
        btnYes.Location = new Point(100, 105);
        btnYes.Size     = new Size(80, 30);
        btnYes.Click   += BtnYes_Click;
        btnYes.Visible  = false;

        // btnNo
        btnNo.Text     = "否";
        btnNo.Location = new Point(200, 105);
        btnNo.Size     = new Size(80, 30);
        btnNo.Click   += BtnNo_Click;
        btnNo.Visible  = false;

        ClientSize      = new Size(484, 151);
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox     = false;
        MinimizeBox     = false;
        StartPosition   = FormStartPosition.CenterParent;
        Controls.AddRange(new Control[] { lblMessage, btnOk, btnCancel, btnYes, btnNo });
        ResumeLayout(false);
    }
}
