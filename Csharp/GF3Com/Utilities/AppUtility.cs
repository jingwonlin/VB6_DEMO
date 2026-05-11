using System.Drawing.Printing;
using System.Text;
using GF3Com.Interop;

namespace GF3Com.Utilities;

/// <summary>
/// 通用工具方法 (from Utility.bas)
/// </summary>
internal static class AppUtility
{
    // ── 訊息框 ───────────────────────────────────────────────────────

    /// <summary>
    /// 顯示自訂訊息框 (對應 xMsgBox in Utility.bas)
    /// autoCloseMs > 0 時會自動關閉（以 Timer 實作）
    /// </summary>
    public static DialogResult ShowMessage(
        string message,
        MessageBoxButtons buttons = MessageBoxButtons.OK,
        MessageBoxIcon icon = MessageBoxIcon.Information,
        string? title = null,
        int autoCloseMs = 0)
    {
        string caption = title ?? Application.ProductName;

        if (autoCloseMs > 0)
        {
            // 自動關閉：用 Task 在背景等待後關閉最上層表單
            using var cts = new CancellationTokenSource(autoCloseMs);
            var token = cts.Token;
            _ = Task.Delay(autoCloseMs, token).ContinueWith(_ =>
            {
                if (Form.ActiveForm is Form f && !f.IsDisposed)
                    f.BeginInvoke(() => f.Close());
            }, TaskScheduler.Default);
        }

        return MessageBox.Show(message, caption, buttons, icon);
    }

    // ── 表單置中 ─────────────────────────────────────────────────────

    /// <summary>
    /// 將表單置中於螢幕 (CenterForm / CentrolForm in Utility.bas)
    /// </summary>
    public static void CenterForm(Form form)
    {
        var screen = Screen.PrimaryScreen?.Bounds ?? Rectangle.Empty;
        int left = Math.Max(0, (screen.Width  - form.Width)  / 2);
        int top  = Math.Max(0, (screen.Height - form.Height) / 2);
        form.Location = new Point(left, top);
    }

    /// <summary>
    /// 置中偏上 (CentrolFormL in Utility.bas)
    /// </summary>
    public static void CenterFormUpper(Form form)
    {
        var screen = Screen.PrimaryScreen?.Bounds ?? Rectangle.Empty;
        int left = Math.Max(0, (screen.Width  - form.Width)  / 2);
        int top  = Math.Max(0, (screen.Height - form.Height) / 2) + 200;
        form.Location = new Point(left, top);
    }

    // ── 狀態列更新 ──────────────────────────────────────────────────

    /// <summary>
    /// 更新主視窗狀態列 (uApp_StatusBar in Utility.bas)
    /// </summary>
    public static void SetStatusBar(StatusStrip? bar, int panelIndex, string text)
    {
        if (bar is null) return;
        if (panelIndex < 1 || panelIndex > bar.Items.Count) return;
        var item = bar.Items[panelIndex - 1] as ToolStripStatusLabel;
        if (item is not null)
            item.Text = text;
    }

    // ── OS 版本 ──────────────────────────────────────────────────────

    /// <summary>
    /// 取得作業系統版本資訊 (apiGetOS_Platform in Utility.bas)
    /// </summary>
    public static (string platform, string version, string build, string servicePack)
        GetOsPlatform()
    {
        var info = Globals.OsVerInfo;
        info.dwOSVersionInfoSize = (uint)System.Runtime.InteropServices.Marshal.SizeOf(info);
        NativeMethods.GetVersionEx(ref info);
        Globals.OsVerInfo = info;

        string platform = info.dwPlatformId switch
        {
            OsVersionInfo.VER_PLATFORM_WIN32_NT when info.dwMajorVersion > 4
                => "Microsoft Windows 2000",
            OsVersionInfo.VER_PLATFORM_WIN32_NT
                => "Microsoft Windows NT",
            OsVersionInfo.VER_PLATFORM_WIN32_WINDOWS when info.dwMinorVersion == 0
                => "Microsoft Windows 95",
            OsVersionInfo.VER_PLATFORM_WIN32_WINDOWS
                => "Microsoft Windows 98",
            OsVersionInfo.VER_PLATFORM_WIN32s
                => "Microsoft Windows 3.1",
            _ => "Unknown"
        };

        string version = $"{info.dwMajorVersion:D}.{info.dwMinorVersion:D2}";
        string build   = info.dwBuildNumber.ToString();
        string sp      = info.szCSDVersion?[..Math.Min(14, info.szCSDVersion.Length)] ?? "";

        return (platform, version, build, sp);
    }

    // ── 字串工具 ─────────────────────────────────────────────────────

    /// <summary>
    /// null/空值轉 "0" (lLv in Utility.bas)
    /// </summary>
    public static string LLv(object? v) =>
        v == null || v == DBNull.Value ? "0" : Convert.ToString(v) ?? "0";

    /// <summary>
    /// null/空值轉 "" (lv in Utility.bas)
    /// </summary>
    public static string Lv(object? v) =>
        v == null || v == DBNull.Value ? string.Empty : Convert.ToString(v) ?? string.Empty;

    /// <summary>
    /// 固定長度字串截斷或填補空白 (VB6 String * N 行為)
    /// </summary>
    public static string FixedLen(string? value, int length) =>
        (value ?? string.Empty).PadRight(length)[..length];

    /// <summary>
    /// 多位元組字元感知的 Mid 函式 (usmid in Utility.bas)
    /// 對應 VB6 中處理 Big5 雙位元組字元的邏輯。
    /// </summary>
    public static string UsMid(string s, int start, int wLen)
    {
        if (string.IsNullOrEmpty(s)) return string.Empty;
        var enc   = Encoding.GetEncoding(950);
        byte[] bs = enc.GetBytes(s);

        int inx = 0;
        var result = new List<byte>();

        for (int i = 0; i < bs.Length; i++)
        {
            inx++;
            // Big5 首位元組：0x81–0xFE
            bool isLeadByte = bs[i] >= 0x81;
            if (isLeadByte && wLen >= 1) inx++;

            if (inx >= start && inx <= start + wLen - 1)
            {
                result.Add(bs[i]);
                if (isLeadByte && i + 1 < bs.Length)
                    result.Add(bs[++i]);
            }
            if (inx > start + wLen - 1) break;
        }
        return enc.GetString(result.ToArray());
    }

    // ── 延遲 ─────────────────────────────────────────────────────────

    /// <summary>
    /// 等待指定秒數同時保持 UI 回應 (delay_sec in Utility.bas)
    /// </summary>
    public static void DelaySec(double seconds)
    {
        var end = DateTime.Now.AddSeconds(seconds);
        while (DateTime.Now < end)
            Application.DoEvents();
    }

    // ── 列印截圖 ─────────────────────────────────────────────────────

    /// <summary>
    /// 將表單截圖列印 (H_Copy in Utility.bas)
    /// </summary>
    public static void PrintFormCapture(Form form, PictureBox picBox,
                                         int rotation = 0)
    {
        // 擷取表單視窗 DC
        IntPtr hDC  = NativeMethods.GetWindowDC(form.Handle);
        int    sw   = form.Width;
        int    sh   = form.Height;

        picBox.Width  = sw + 4;
        picBox.Height = sh + 4;

        using var bmp = new Bitmap(picBox.Width, picBox.Height);
        using var g   = Graphics.FromImage(bmp);
        IntPtr hDest  = g.GetHdc();

        NativeMethods.BitBlt(hDest, 0, 0, sw + 4, sh + 4,
                             hDC, 0, 0, NativeMethods.SRCCOPY);
        g.ReleaseHdc(hDest);
        NativeMethods.ReleaseDC(form.Handle, hDC);

        picBox.Image = bmp;

        // 列印
        var pd = new PrintDocument();
        pd.PrintPage += (_, e) =>
        {
            if (e.Graphics is null) return;
            float pw = e.PageBounds.Width;
            float ph = e.PageBounds.Height;
            float iw = bmp.Width;
            float ih = bmp.Height;

            if (iw > pw || ih > ph)
            {
                float scale = Math.Min(pw / iw, ph / ih) * 0.9f;
                iw *= scale;
                ih *= scale;
            }
            float x = (pw - iw) / 2;
            float y = (ph - ih) / 2;
            e.Graphics.DrawImage(bmp, x, y, iw, ih);
        };
        pd.Print();
    }

    // ── 電腦名稱 ─────────────────────────────────────────────────────

    public static string GetComputerName()
    {
        var sb   = new StringBuilder(256);
        uint len = 256;
        NativeMethods.GetComputerName(sb, ref len);
        return sb.ToString();
    }
}
