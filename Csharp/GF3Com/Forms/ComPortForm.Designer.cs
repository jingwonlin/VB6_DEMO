namespace GF3Com.Forms;

partial class ComPortForm
{
    private System.ComponentModel.IContainer components = null;
    private Button cmdEnd = null!;

    protected override void Dispose(bool disposing)
    {
        if (disposing && components != null)
            components.Dispose();
        base.Dispose(disposing);
    }

    private void InitializeComponent()
    {
        cmdEnd = new Button();
        SuspendLayout();

        // 建立各 PLC port 顯示行 (index 21–25)
        string[] frameLabels =
        {
            "熱壓通訊埠",   // 21 = COM4
            "結束通訊埠",   // 22 = COM5
            "冷壓通訊埠",   // 23 = COM6
            "成品通訊埠",   // 24 = COM7
            "原料通訊埠",   // 25 = COM8
        };

        var font = new Font("新細明體", 12f);
        int yBase = 10;

        for (int i = 0; i < 5; i++)
        {
            int idx = i + 21;
            int y   = yBase + i * 110;

            var grp = new GroupBox
            {
                Text     = frameLabels[i],
                Location = new Point(10, y),
                Size     = new Size(960, 100),
                Font     = new Font("新宋體", 12f)
            };

            var lblStrTitle = new Label { Text = "接收字串", Location = new Point(5, 30), Size = new Size(80, 30), Font = font };
            var lblString   = new Label { BorderStyle = BorderStyle.FixedSingle, Location = new Point(90, 30),  Size = new Size(500, 28), ForeColor = Color.Red, Font = font };
            var lblLenTitle = new Label { Text = "Byte",     Location = new Point(600, 30), Size = new Size(50, 28), Font = font };
            var lblLen      = new Label { BorderStyle = BorderStyle.FixedSingle, Location = new Point(650, 30),  Size = new Size(60, 28),  ForeColor = Color.Red, Font = font, TextAlign = ContentAlignment.MiddleCenter };
            var lblStTitle  = new Label { Text = "通訊狀態", Location = new Point(5, 65),  Size = new Size(80, 28), Font = font };
            var lblStatus   = new Label { BorderStyle = BorderStyle.FixedSingle, Location = new Point(90, 65),   Size = new Size(500, 28), ForeColor = Color.Red, Font = font };

            _lblString[idx] = lblString;
            _lblLen[idx]    = lblLen;
            _lblStatus[idx] = lblStatus;

            grp.Controls.AddRange(new Control[]
                { lblStrTitle, lblString, lblLenTitle, lblLen, lblStTitle, lblStatus });
            Controls.Add(grp);
        }

        yBase += 5 * 110 + 10;

        // AGV 群組
        var agvGrp = new GroupBox
        {
            Text     = "AGV通訊埠",
            Location = new Point(10, yBase),
            Size     = new Size(960, 150),
            Font     = new Font("新宋體", 12f)
        };
        var lblSendTitle = new Label { Text = "傳送字串", Location = new Point(5, 30),  Size = new Size(80, 28), Font = new Font("新細明體", 12f) };
        _lblOutString[9] = new Label { BorderStyle = BorderStyle.FixedSingle, Location = new Point(90, 30), Size = new Size(500, 28), ForeColor = Color.Red, Font = new Font("新細明體", 12f) };
        _lblOutLen[9]    = new Label { BorderStyle = BorderStyle.FixedSingle, Location = new Point(600, 30), Size = new Size(60, 28), ForeColor = Color.Red, Font = new Font("新細明體", 12f), TextAlign = ContentAlignment.MiddleCenter };

        var lblAgv1Title = new Label { Text = "AGV1", Location = new Point(30, 70), Size = new Size(60, 25), Font = new Font("新宋體", 12f) };
        _lblAgvStatus[1] = new Label { BorderStyle = BorderStyle.FixedSingle, Location = new Point(100, 70), Size = new Size(400, 25), Font = new Font("新宋體", 12f) };
        var lblAgv2Title = new Label { Text = "AGV2", Location = new Point(30, 105), Size = new Size(60, 25), Font = new Font("新宋體", 12f) };
        _lblAgvStatus[2] = new Label { BorderStyle = BorderStyle.FixedSingle, Location = new Point(100, 105), Size = new Size(400, 25), Font = new Font("新宋體", 12f) };

        agvGrp.Controls.AddRange(new Control[]
        {
            lblSendTitle, _lblOutString[9], _lblOutLen[9],
            lblAgv1Title, _lblAgvStatus[1],
            lblAgv2Title, _lblAgvStatus[2]
        });
        Controls.Add(agvGrp);

        // 關閉按鈕
        cmdEnd.Text     = "關閉";
        cmdEnd.Location = new Point(880, yBase + 160);
        cmdEnd.Size     = new Size(90, 40);
        cmdEnd.Font     = new Font("新宋體", 12f);
        cmdEnd.Click   += CmdEnd_Click;
        Controls.Add(cmdEnd);

        ClientSize      = new Size(980, yBase + 215);
        FormBorderStyle = FormBorderStyle.Sizable;
        Text            = "ComPort";
        WindowState     = FormWindowState.Maximized;
        StartPosition   = FormStartPosition.WindowsDefaultLocation;
        ResumeLayout(false);
    }
}
