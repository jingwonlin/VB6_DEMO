namespace GF3Com.Constants;

/// <summary>
/// MOXA api232.dll 鮑率 / 資料位元 / 停止位元 / 同位元常數 (from Api232-b.bas)
/// </summary>
internal static class SerialConstants
{
    // Baud Rate
    public const int B50     = 0x00;
    public const int B75     = 0x01;
    public const int B110    = 0x02;
    public const int B134    = 0x03;
    public const int B150    = 0x04;
    public const int B300    = 0x05;
    public const int B600    = 0x06;
    public const int B1200   = 0x07;
    public const int B1800   = 0x08;
    public const int B2400   = 0x09;
    public const int B4800   = 0x0A;
    public const int B7200   = 0x0B;
    public const int B9600   = 0x0C;
    public const int B19200  = 0x0D;
    public const int B38400  = 0x0E;
    public const int B57600  = 0x0F;
    public const int B115200 = 0x10;
    public const int B230400 = 0x11;
    public const int B460800 = 0x12;
    public const int B921600 = 0x13;

    // Word length (data bits)
    public const int BIT_5 = 0x00;
    public const int BIT_6 = 0x01;
    public const int BIT_7 = 0x02;
    public const int BIT_8 = 0x03;

    // Stop bits
    public const int STOP_1 = 0x00;
    public const int STOP_2 = 0x04;

    // Parity
    public const int P_EVEN = 0x18;
    public const int P_ODD  = 0x08;
    public const int P_SPC  = 0x38;
    public const int P_MRK  = 0x28;
    public const int P_NONE = 0x00;

    // Modem control
    public const int C_DTR = 0x01;
    public const int C_RTS = 0x02;

    // Modem line status
    public const int S_CTS = 0x01;
    public const int S_DSR = 0x02;
    public const int S_RI  = 0x04;
    public const int S_CD  = 0x08;

    // Error codes
    public const int SIO_OK         =  0;
    public const int SIO_BADPORT    = -1;
    public const int SIO_OUTCONTROL = -2;
    public const int SIO_NODATA     = -4;
    public const int SIO_BADPARM    = -7;
    public const int SIO_WIN32FAIL  = -8;
}
