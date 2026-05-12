using GF3Com.Forms;

// 繁體中文（Big5）支援
System.Text.Encoding.RegisterProvider(
    System.Text.CodePagesEncodingProvider.Instance);

Application.EnableVisualStyles();
Application.SetCompatibleTextRenderingDefault(false);
Application.SetHighDpiMode(HighDpiMode.SystemAware);

Application.Run(new MainForm());
