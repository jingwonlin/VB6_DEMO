using System.IO.Ports;
using GF3Com.Interop;

namespace GF3Com;

/// <summary>
/// 全域狀態 (from Global.bas + SK900_plc.bas)
/// VB6 模組層級 Public 變數統一集中於此靜態類別。
/// </summary>
internal static class Globals
{
    // ── 資料庫連線字串 ───────────────────────────────────────────────
    public static string SqlConnectionString { get; set; } = string.Empty;
    public static string SqlConnectionStringLocal { get; set; } = string.Empty;

    // ── 系統識別 ────────────────────────────────────────────────────
    public static string ComputerName  { get; set; } = string.Empty;
    public static string ProjectName   { get; set; } = string.Empty;   // SK01 / SK02
    public static string PlantName     { get; set; } = string.Empty;   // 南亞二 / 南亞三
    public static string DbType        { get; set; } = "SQL";

    // ── 使用者資訊 ──────────────────────────────────────────────────
    public static string LoginUser { get; set; } = new string(' ', 10);
    public static string LoginName { get; set; } = new string(' ', 8);
    public static string LoginPass { get; set; } = new string(' ', 5);
    public static string Super     { get; set; } = "N";
    public static string Awno      { get; set; } = new string(' ', 1);
    public static string Dept      { get; set; } = new string(' ', 10);

    // ── AGV 通訊 ────────────────────────────────────────────────────
    public static int    AgvPort      { get; set; } = 9;
    public static int    AgvOpenOk    { get; set; } = 0;
    public static string AgvTempStr  { get; set; } = string.Empty;
    public static byte[] SendAgvStr  { get; set; } = Array.Empty<byte>();
    public static byte   Bcc         { get; set; } = 0;

    // ── PLC 狀態緩衝 ────────────────────────────────────────────────
    // 熱壓 (ST21, COM4+17=21)
    public static string[] ST21_Temp  = new string[4];   // index 1–3
    public static string[] ST21_StNo  = new string[4];
    public static int[]    ST21_StByte= new int[4];

    // 結束 (ST22, COM5+17=22)
    public static string[] ST22_Temp  = new string[8];   // index 1–7
    public static string[] ST22_StNo  = new string[8];
    public static int[]    ST22_StByte= new int[8];

    // 冷壓 (ST23, COM6+17=23)
    public static string[] ST23_Temp  = new string[7];   // index 1–6
    public static string[] ST23_StNo  = new string[7];
    public static int[]    ST23_StByte= new int[7];

    // 成品 (ST24, COM7+17=24)
    public static string[] ST24_Temp  = new string[6];   // index 1–5
    public static string[] ST24_StNo  = new string[6];
    public static int[]    ST24_StByte= new int[6];

    // 原料 (ST25, COM8+17=25)
    public static string[] ST25_Temp  = new string[12];  // index 1–11
    public static string[] ST25_StNo  = new string[12];
    public static int[]    ST25_StByte= new int[12];

    // PLC 回應期望長度 (index 21–25)
    public static readonly int[] PlcLen = new int[26];

    // ── OS 版本 ─────────────────────────────────────────────────────
    public static OsVersionInfo OsVerInfo;

    // ── 系統旗標 ────────────────────────────────────────────────────
    public static bool  CranMoveFlag  { get; set; } = true;
    public static int   Refresh1Screen{ get; set; } = 0;
    public static int   Plc1Num       { get; set; } = 7;
    public static int   StnoNum       { get; set; } = 5;
    public static int   Err2Cnt       { get; set; } = 0;
    public static string BackupDir    { get; set; } = @"c:\";
    public static string Err1Place    { get; set; } = string.Empty;

    // ── 各天車錯誤號碼 (index 1–5) ──────────────────────────────────
    public static readonly int[] CraneErrNo  = new int[6];
    public static readonly int[] CraneTemp   = new int[6];

    // ── COM Port 開關狀態 (index 4–8) ───────────────────────────────
    public static readonly int[] ComOpenOk = new int[10];

    // ── AGV 命令計數 ────────────────────────────────────────────────
    public static int AgvCommandCnt { get; set; } = 0;

    /// <summary>初始化所有陣列預設值（對應 InitialParameter）</summary>
    public static void InitialParameter()
    {
        for (int i = 1; i <= 5; i++)
            CraneErrNo[i] = 99;

        for (int i = 1; i <= 3; i++)  ST21_Temp[i] = "9";
        for (int i = 1; i <= 7; i++)  ST22_Temp[i] = "9";
        for (int i = 1; i <= 6; i++)  ST23_Temp[i] = "9";
        for (int i = 1; i <= 5; i++)  ST24_Temp[i] = "9";
        for (int i = 1; i <= 9; i++)  ST25_Temp[i] = "9";

        PlcLen[21] = 28;   // 熱壓
        PlcLen[22] = 20;   // 結束
        PlcLen[23] = 14;   // 冷壓
        PlcLen[24] = 16;   // 成品
        PlcLen[25] = 50;   // 原料

        ST21_StNo[1] = "A003"; ST21_StNo[2] = "A004"; ST21_StNo[3] = "A005";
        ST21_StByte[1] = 7;    ST21_StByte[2] = 9;    ST21_StByte[3] = 11;

        ST22_StNo[1] = "B001"; ST22_StNo[2] = "B002"; ST22_StNo[3] = "B003";
        ST22_StNo[4] = "B004"; ST22_StNo[5] = "B005"; ST22_StNo[6] = "B006"; ST22_StNo[7] = "B007";
        ST22_StByte[1] = 3; ST22_StByte[2] = 4; ST22_StByte[3] = 5;
        ST22_StByte[4] = 6; ST22_StByte[5] = 7; ST22_StByte[6] = 8; ST22_StByte[7] = 9;

        ST23_StNo[1] = "C001"; ST23_StNo[2] = "C002"; ST23_StNo[3] = "C003";
        ST23_StNo[4] = "C004"; ST23_StNo[5] = "C005"; ST23_StNo[6] = "C006";
        ST23_StByte[1] = 3; ST23_StByte[2] = 4; ST23_StByte[3] = 5;
        ST23_StByte[4] = 6; ST23_StByte[5] = 7; ST23_StByte[6] = 8;

        ST24_StNo[1] = "D001"; ST24_StNo[2] = "D002"; ST24_StNo[3] = "D003";
        ST24_StNo[4] = "D004"; ST24_StNo[5] = "D005";
        ST24_StByte[1] = 3; ST24_StByte[2] = 4; ST24_StByte[3] = 5;
        ST24_StByte[4] = 6; ST24_StByte[5] = 7;

        ST25_StNo[1] = "E001"; ST25_StNo[2]  = "E002"; ST25_StNo[3]  = "E003";
        ST25_StNo[4] = "E004"; ST25_StNo[5]  = "E005"; ST25_StNo[6]  = "E006";
        ST25_StNo[7] = "E007"; ST25_StNo[8]  = "E008"; ST25_StNo[9]  = "E009";
        ST25_StNo[10]= "E010";
        ST25_StByte[1]=5;  ST25_StByte[2]=4;  ST25_StByte[3]=8;  ST25_StByte[4]=7;
        ST25_StByte[5]=6;  ST25_StByte[6]=10; ST25_StByte[7]=9;  ST25_StByte[8]=12;
        ST25_StByte[9]=11; ST25_StByte[10]=13;
    }
}
