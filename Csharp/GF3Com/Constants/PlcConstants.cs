namespace GF3Com.Constants;

/// <summary>
/// PLC / 設備狀態常數 (from SK900_plc.bas)
/// </summary>
internal static class PlcConstants
{
    // 設備名稱
    public const string EQ_NULL    = "FINISH    ";
    public const string EQ_STV1    = "STV1      ";
    public const string EQ_STV2    = "STV2      ";
    public const string EQ_MONO1   = "MONO1     ";
    public const string EQ_LIFTER1 = "LIFTER1   ";
    public const string EQ_ROTATE  = "ROTATE    ";

    // 儲格狀態
    public const string SG_EMPTY  = "0";
    public const string SG_LKIN   = "S";
    public const string SG_LKOUT  = "R";
    public const string SG_STOR   = "F";
    public const string SG_ERR    = "E";
    public const string SG_NODATA = "N";

    // 移動狀態
    public const string MV_NONE = "0";
    public const string MV_NOW  = "1";
    public const string MV_ERR1 = "2";
    public const string MV_ERR  = "E";

    public const string NOLD    = "0";
    public const string FB_NONE = "0";
    public const string FB_ALL  = "X";

    // 站台狀態
    public const string ST_RQST = "0";
    public const string ST_WAIT = "1";
    public const string ST_WORK = "2";

    // 交易類型
    public const string TR_IN          = "S";
    public const string TR_OUT         = "R";
    public const string TR_NORMAL_OUT  = "1";
    public const string TR_PALLETS_OUT = "2";
    public const string TR_QOUT        = "3";
    public const string TR_MOVE_OUT    = "4";
    public const string TR_NO_OUT      = "5";
    public const string TR_NORMAL_IN   = "A";
    public const string TR_PALLETS_IN  = "B";
    public const string TR_QIN         = "C";
    public const string TR_MOVE_IN     = "D";
    public const string TR_BAR_IN      = "E";
    public const string TR_OUT_REIN    = "F";
    public const string TR_MIX_IN      = "G";
    public const string TR_MOVE_STNO   = "I";
    public const string TR_MIN         = "K";
    public const string TR_MOUT        = "L";
    public const string TR_MOD         = "M";
    public const string TR_MOD1        = "N";
    public const string TR_3000        = "O";
    public const string TR_3002        = "P";
    public const string TR_CMP         = "Q";
    public const string TR_DLT         = "R";
    public const string TR_RESEND      = "S";
    public const string TR_FULL        = "V";
    public const string TR_EMPTY       = "W";

    // PLC 回應長度 (port index 21–25 對應 COM4–COM8)
    public static readonly IReadOnlyDictionary<int, int> PlcResponseLength = new Dictionary<int, int>
    {
        { 21, 28 },   // 熱壓
        { 22, 20 },   // 結束
        { 23, 14 },   // 冷壓
        { 24, 16 },   // 成品
        { 25, 50 },   // 原料
    };

    // COM port 對應的 PLC 查詢命令
    public static string GetPlcQuery(int portIndex) => portIndex switch
    {
        4 => ">3\r\n",   // 熱壓
        5 => ">2\r\n",   // 結束
        6 => ">2\r\n",   // 冷壓
        7 => ">2\r\n",   // 成品
        8 => ">5\r\n",   // 原料
        _ => string.Empty
    };
}
