using System.Text;
using GF3Com.Interop;

namespace GF3Com.Utilities;

/// <summary>
/// INI 設定檔讀取 (from IniRead in Utility.bas)
/// 保留與原始 GF3Com.ini 的相容性。
/// </summary>
internal static class IniReader
{
    private const int BufferSize = 1024;

    /// <summary>
    /// 讀取 INI 設定值 (對應 IniRead function in Utility.bas)
    /// </summary>
    /// <param name="section">區段名稱 (lpAppName)</param>
    /// <param name="key">鍵名 (lpKeyName)；傳入 "NULL" 可列舉所有鍵</param>
    /// <param name="filePath">INI 檔案完整路徑</param>
    /// <returns>設定值；找不到時回傳 string.Empty</returns>
    public static string Read(string section, string key, string filePath)
    {
        var sb      = new StringBuilder(BufferSize);
        string def  = "KeyNull";
        string? actualKey = string.Equals(key.Trim(), "NULL",
            StringComparison.OrdinalIgnoreCase) ? null : key;

        uint written = NativeMethods.GetPrivateProfileString(
            section, actualKey, def, sb, (uint)BufferSize, filePath);

        string result = sb.ToString(0, (int)written);

        if (result == "KeyNull")
        {
            AppUtility.ShowMessage(
                $"{filePath} 找不到 [{section}] {key} = 的描述",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            return string.Empty;
        }
        return result;
    }

    /// <summary>
    /// 安全讀取：找不到時回傳 defaultValue，不顯示錯誤對話框
    /// </summary>
    public static string ReadOrDefault(string section, string key,
                                        string filePath, string defaultValue = "")
    {
        var sb = new StringBuilder(BufferSize);
        NativeMethods.GetPrivateProfileString(section, key, defaultValue, sb,
            (uint)BufferSize, filePath);
        return sb.ToString();
    }
}
