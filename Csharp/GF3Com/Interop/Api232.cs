using System.Runtime.InteropServices;
using System.Text;
using GF3Com.Constants;

namespace GF3Com.Interop;

/// <summary>
/// MOXA api232.dll P/Invoke 宣告 (from Api232-b.bas)
/// 注意：api232.dll 為 32-bit，專案必須以 x86 平台編譯。
/// 若改用 System.IO.Ports.SerialPort 可移除此類別。
/// </summary>
internal static class Api232
{
    private const string DllName = "api232.dll";

    [DllImport(DllName, CallingConvention = CallingConvention.Cdecl)]
    internal static extern int sio_open(int port);

    [DllImport(DllName, CallingConvention = CallingConvention.Cdecl)]
    internal static extern int sio_close(int port);

    [DllImport(DllName, CallingConvention = CallingConvention.Cdecl)]
    internal static extern int sio_ioctl(int port, int baud, int mode);

    [DllImport(DllName, CallingConvention = CallingConvention.Cdecl,
        CharSet = CharSet.Ansi)]
    internal static extern int sio_read(int port, StringBuilder buf, int length);

    [DllImport(DllName, CallingConvention = CallingConvention.Cdecl,
        CharSet = CharSet.Ansi)]
    internal static extern int sio_write(int port, string buf, int length);

    [DllImport(DllName, CallingConvention = CallingConvention.Cdecl)]
    internal static extern int sio_putch(int port, int term);

    [DllImport(DllName, CallingConvention = CallingConvention.Cdecl)]
    internal static extern int sio_getch(int port);

    [DllImport(DllName, CallingConvention = CallingConvention.Cdecl)]
    internal static extern int sio_flush(int port, int func);

    [DllImport(DllName, CallingConvention = CallingConvention.Cdecl)]
    internal static extern int sio_iqueue(int port);

    [DllImport(DllName, CallingConvention = CallingConvention.Cdecl)]
    internal static extern int sio_oqueue(int port);

    [DllImport(DllName, CallingConvention = CallingConvention.Cdecl)]
    internal static extern int sio_lstatus(int port);

    [DllImport(DllName, CallingConvention = CallingConvention.Cdecl)]
    internal static extern int sio_lctrl(int port, int mode);

    [DllImport(DllName, CallingConvention = CallingConvention.Cdecl)]
    internal static extern int sio_flowctrl(int port, int mode);

    [DllImport(DllName, CallingConvention = CallingConvention.Cdecl)]
    internal static extern int sio_DTR(int port, int mode);

    [DllImport(DllName, CallingConvention = CallingConvention.Cdecl)]
    internal static extern int sio_RTS(int port, int mode);

    [DllImport(DllName, CallingConvention = CallingConvention.Cdecl)]
    internal static extern int sio_baud(int port, int speed);

    [DllImport(DllName, CallingConvention = CallingConvention.Cdecl)]
    internal static extern int sio_data_status(int port);

    [DllImport(DllName, CallingConvention = CallingConvention.Cdecl)]
    internal static extern int sio_SetReadTimeouts(int port, int totalTimeouts, int intervalTimeouts);

    [DllImport(DllName, CallingConvention = CallingConvention.Cdecl)]
    internal static extern int sio_AbortRead(int port);

    [DllImport(DllName, CallingConvention = CallingConvention.Cdecl)]
    internal static extern int sio_SetWriteTimeouts(int port, int timeouts);

    [DllImport(DllName, CallingConvention = CallingConvention.Cdecl)]
    internal static extern int sio_AbortWrite(int port);

    /// <summary>
    /// 開啟並初始化 MOXA 串列埠 (from moxaOpen in Api232-b.bas)
    /// </summary>
    internal static int MoxaOpen(int comPort, int baud, string parity, int dataBits, int stopBits)
    {
        int result = sio_open(comPort);
        if (result < 0)
            return result;

        int baudCode = baud switch
        {
            2400  => SerialConstants.B2400,
            4800  => SerialConstants.B4800,
            9600  => SerialConstants.B9600,
            19200 => SerialConstants.B19200,
            38400 => SerialConstants.B38400,
            _     => -2
        };
        if (baudCode == -2) return -2;

        int bitCode = dataBits switch
        {
            8 => SerialConstants.BIT_8,
            7 => SerialConstants.BIT_7,
            _ => -2
        };
        if (bitCode == -2) return -2;

        int stopCode = stopBits switch
        {
            1 => SerialConstants.STOP_1,
            2 => SerialConstants.STOP_2,
            _ => -2
        };
        if (stopCode == -2) return -2;

        int parityCode = parity.ToUpperInvariant() switch
        {
            "EVEN" => SerialConstants.P_EVEN,
            "ODD"  => SerialConstants.P_ODD,
            "NONE" => SerialConstants.P_NONE,
            _      => -2
        };
        if (parityCode == -2) return -2;

        return sio_ioctl(comPort, baudCode, bitCode | stopCode | parityCode);
    }
}
