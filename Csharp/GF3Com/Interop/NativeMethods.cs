using System.Runtime.InteropServices;
using System.Text;

namespace GF3Com.Interop;

/// <summary>
/// Win32 API P/Invoke 宣告 (from Global.bas)
/// </summary>
internal static class NativeMethods
{
    // SetWindowPos flags
    public const uint SWP_NOSIZE   = 0x0001;
    public const uint SWP_NOMOVE   = 0x0002;
    public const uint SWP_NOZORDER = 0x0004;
    public const uint SWP_SHOWWINDOW = 0x0040;

    // BitBlt raster operation
    public const uint SRCCOPY = 0x00CC0020;

    [DllImport("user32.dll", SetLastError = true)]
    internal static extern bool SetWindowPos(
        IntPtr hWnd,
        IntPtr hWndInsertAfter,
        int x, int y, int cx, int cy,
        uint uFlags);

    [DllImport("kernel32.dll", CharSet = CharSet.Ansi, SetLastError = true)]
    internal static extern uint GetPrivateProfileString(
        string lpAppName,
        string? lpKeyName,
        string lpDefault,
        StringBuilder lpReturnedString,
        uint nSize,
        string lpFileName);

    [DllImport("kernel32.dll", CharSet = CharSet.Ansi)]
    internal static extern bool GetVersionEx(ref OsVersionInfo lpVersionInformation);

    [DllImport("user32.dll")]
    internal static extern IntPtr GetWindowDC(IntPtr hWnd);

    [DllImport("gdi32.dll")]
    internal static extern bool BitBlt(
        IntPtr hdcDest, int nXDest, int nYDest, int nWidth, int nHeight,
        IntPtr hdcSrc,  int nXSrc,  int nYSrc,
        uint dwRop);

    [DllImport("user32.dll")]
    internal static extern int ReleaseDC(IntPtr hWnd, IntPtr hDC);

    [DllImport("kernel32.dll")]
    internal static extern void Sleep(uint dwMilliseconds);

    [DllImport("kernel32.dll", CharSet = CharSet.Ansi)]
    internal static extern bool GetComputerName(StringBuilder lpBuffer, ref uint nSize);
}

/// <summary>
/// OSVERSIONINFO 結構 (from Global.bas)
/// </summary>
[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
internal struct OsVersionInfo
{
    public uint dwOSVersionInfoSize;
    public uint dwMajorVersion;
    public uint dwMinorVersion;
    public uint dwBuildNumber;
    public uint dwPlatformId;
    [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 128)]
    public string szCSDVersion;

    public const uint VER_PLATFORM_WIN32s        = 0;
    public const uint VER_PLATFORM_WIN32_WINDOWS = 1;
    public const uint VER_PLATFORM_WIN32_NT      = 2;
}
