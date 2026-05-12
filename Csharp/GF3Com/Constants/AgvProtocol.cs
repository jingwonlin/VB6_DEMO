namespace GF3Com.Constants;

/// <summary>
/// AGV 通訊協定常數 (from Global.bas / Api232-b.bas)
/// </summary>
internal static class AgvProtocol
{
    public const string HOST  = "1";
    public const string AGV   = "2";
    public const string DUMMY = "0";
    public const string READY = "1";

    public const string GET_EMPTY = "1";
    public const string PUT_EMPTY = "2";
    public const string GET_FULL  = "3";
    public const string PUT_FULL  = "4";

    // 站點代碼
    public const string D001 = "C1   ";
    public const string D002 = "C2   ";
    public const string D003 = "C3   ";
    public const string D004 = "C4   ";

    // 控制位元組
    public const byte CODE = 0xFF;
    public const byte DLE  = 0x10;
    public const byte STX  = 0x02;
    public const byte ETX  = 0x03;
    public const byte ACK  = 0x06;
    public const byte NAK  = 0x15;
    public const byte ENQ  = 0x05;
    public const byte EOT  = 0x04;
}
