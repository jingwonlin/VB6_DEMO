# VB6 → C# 轉換分析報告

**專案**: GF3Com (南亞科技倉儲自動化通訊系統)  
**版本**: 1.0.45  
**分析日期**: 2026-05-11

---

## 一、專案概述

GF3Com 是一套用於控制 AGV（自動導引車）與 PLC（可程式邏輯控制器）的倉儲自動化通訊系統，透過多個 RS-232 串列埠（COM4–COM9）進行裝置通訊，並連接 ADODB SQL 資料庫。

### 原始檔案清單

| 檔案 | 類型 | 說明 |
|------|------|------|
| `Global.bas` | 模組 | 全域變數、Win32 API 宣告、ADODB 物件 |
| `SK900_plc.bas` | 模組 | PLC 通訊邏輯（約 355 KB，為最大檔案） |
| `Utility.bas` | 模組 | 工具函式（INI 讀取、訊息框、畫面置中、列印） |
| `Api232-b.bas` | 模組 | MOXA api232.dll 串列埠 API 包裝 |
| `MDImainForm.frm` | 表單 | 主 MDI 視窗（MSComm 控制項、Timer、StatusBar） |
| `frmComPort.frm` | 表單 | COM Port 監控介面 |
| `frmMsgBox.frm` | 表單 | 自訂訊息框 |
| `frmStart.frm` | 表單 | 啟動畫面 |
| `frmApStartMsg.frm` | 表單 | 應用程式啟動訊息 |

---

## 二、轉換難點分析

### 2.1 Win32 API 呼叫

**原始 VB6 (`Global.bas`)**
```vb
Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ...) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (...) As Long
Declare Function BitBlt Lib "gdi32" (...) As Long
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
```

**C# 對應方式**
```csharp
using System.Runtime.InteropServices;

internal static class NativeMethods
{
    [DllImport("user32.dll")]
    internal static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter,
        int x, int y, int cx, int cy, uint uFlags);

    [DllImport("kernel32.dll", CharSet = CharSet.Ansi)]
    internal static extern uint GetPrivateProfileString(string lpAppName,
        string lpKeyName, string lpDefault, StringBuilder lpReturnedString,
        uint nSize, string lpFileName);

    [DllImport("gdi32.dll")]
    internal static extern bool BitBlt(IntPtr hdcDest, int nXDest, int nYDest,
        int nWidth, int nHeight, IntPtr hdcSrc, int nXSrc, int nYSrc, uint dwRop);
}
```

> **建議**：`GetPrivateProfileString` 可改用 .NET 的 `System.Configuration` 或 JSON/XML 設定檔取代 INI 檔案，`Sleep` 改用 `Thread.Sleep()` 或 `await Task.Delay()`。

---

### 2.2 ADODB 資料庫連線

**原始 VB6 (`Global.bas`, `SK900_plc.bas`)**
```vb
Public gADOCon As ADODB.Connection
Public gADOCmd As ADODB.Command
Public gADORecord As New ADODB.Recordset
Public rcn As New ADODB.Connection, rcn1 As New ADODB.Connection, ...
```

**C# 對應方式**

改用 `System.Data.SqlClient`（或 `Microsoft.Data.SqlClient`）：
```csharp
using Microsoft.Data.SqlClient;

public class DatabaseContext
{
    private readonly string _connectionString;
    private SqlConnection _connection;

    public void Connect(string connStr)
    {
        _connection = new SqlConnection(connStr);
        _connection.Open();
    }

    public DataTable ExecuteQuery(string sql)
    {
        using var adapter = new SqlDataAdapter(sql, _connection);
        var dt = new DataTable();
        adapter.Fill(dt);
        return dt;
    }
}
```

> **建議**：若需 ORM 可考慮 Entity Framework Core；若需保持類似 ADO 的操作方式，`SqlConnection`/`SqlCommand`/`SqlDataReader` 是最直接的對應。

---

### 2.3 MOXA api232.dll 串列埠通訊

**原始 VB6 (`Api232-b.bas`)**
```vb
Declare Function sio_open  Lib "api232.dll" (ByVal port As Long) As Long
Declare Function sio_ioctl Lib "api232.dll" (ByVal port As Long, ByVal Baud As Long, ByVal mode As Long) As Long
Declare Function sio_read  Lib "api232.dll" (ByVal port As Long, ByVal buf As String, ByVal length As Long) As Long
Declare Function sio_write Lib "api232.dll" (ByVal port As Long, ByVal buf As String, ByVal length As Long) As Long
Declare Function sio_close Lib "api232.dll" (ByVal port As Long) As Long
```

**C# 有兩種選擇：**

**方案 A — 繼續使用 api232.dll（P/Invoke）**
```csharp
internal static class Api232
{
    [DllImport("api232.dll", CallingConvention = CallingConvention.Cdecl)]
    internal static extern int sio_open(int port);

    [DllImport("api232.dll", CallingConvention = CallingConvention.Cdecl)]
    internal static extern int sio_ioctl(int port, int baud, int mode);

    [DllImport("api232.dll", CallingConvention = CallingConvention.Cdecl,
        CharSet = CharSet.Ansi)]
    internal static extern int sio_read(int port, StringBuilder buf, int length);

    [DllImport("api232.dll", CallingConvention = CallingConvention.Cdecl,
        CharSet = CharSet.Ansi)]
    internal static extern int sio_write(int port, string buf, int length);

    [DllImport("api232.dll", CallingConvention = CallingConvention.Cdecl)]
    internal static extern int sio_close(int port);
}
```

**方案 B — 改用 `System.IO.Ports.SerialPort`（建議）**
```csharp
using System.IO.Ports;

public class SerialPortWrapper : IDisposable
{
    private SerialPort _port;

    public void Open(int comNumber, int baud, Parity parity, int dataBits, StopBits stopBits)
    {
        _port = new SerialPort($"COM{comNumber}", baud, parity, dataBits, stopBits);
        _port.ReceivedBytesThreshold = 1;
        _port.DataReceived += OnDataReceived;
        _port.Open();
    }

    private void OnDataReceived(object sender, SerialDataReceivedEventArgs e)
    {
        int bytesAvailable = _port.BytesToRead;
        byte[] buffer = new byte[bytesAvailable];
        _port.Read(buffer, 0, bytesAvailable);
        // 處理接收資料
    }

    public void Write(byte[] data) => _port.Write(data, 0, data.Length);

    public void Dispose() => _port?.Dispose();
}
```

> **建議採用方案 B**，因為 `System.IO.Ports.SerialPort` 是原生 .NET API，不依賴 32-bit DLL，未來維護性更好。若 MOXA 卡有特殊功能（如多埠擴充），再考慮方案 A。

---

### 2.4 MSComm 控制項（COM4–COM8）

**原始 VB6 (`MDImainForm.frm`)**
```vb
Begin MSCommLib.MSComm MSCommAR
    Index    = 4
    CommPort = 4
    RThreshold = 28
End
```

**C# 對應（`System.IO.Ports.SerialPort`）**
```csharp
private SerialPort[] _commPorts = new SerialPort[9]; // 索引 4–8

private void InitPorts()
{
    int[] rThresholds = { 0, 0, 0, 0, 28, 16, 12, 16, 50 };
    for (int i = 4; i <= 8; i++)
    {
        _commPorts[i] = new SerialPort($"COM{i}", 9600, Parity.Even, 8, StopBits.One);
        _commPorts[i].ReceivedBytesThreshold = rThresholds[i];
        _commPorts[i].DataReceived += (s, e) => OnCommDataReceived(i, s, e);
        _commPorts[i].Open();
    }
}
```

---

### 2.5 MDI 表單架構

**原始 VB6**：`VB.MDIForm` + 子 `VB.Form`

**C# WinForms 對應**
```csharp
// 主 MDI 容器
public partial class MainForm : Form
{
    public MainForm()
    {
        InitializeComponent();
        this.IsMdiContainer = true;
        this.WindowState = FormWindowState.Maximized;
    }
}

// 子表單
public partial class ComPortForm : Form
{
    public ComPortForm()
    {
        InitializeComponent();
        this.MdiParent = Application.OpenForms["MainForm"];
    }
}
```

---

### 2.6 固定長度字串

**原始 VB6 (`SK900_plc.bas`)**
```vb
Public tty As String * 60
Public loginuser As String * 10
Public plc_buf As String * 2000
```

**C# 對應**（無原生固定長度字串，可用 `struct` + `MarshalAs` 或填充方法）
```csharp
// 方案 A：用 struct 模擬（與非受控程式碼互動時）
[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
struct FixedStrings
{
    [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 60)]
    public string tty;
    [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 10)]
    public string loginuser;
}

// 方案 B：直接用 string，在賦值時截斷或填補
public static string FixedLength(string value, int length) =>
    (value ?? "").PadRight(length).Substring(0, length);
```

---

### 2.7 BCC 計算與通訊協定

**原始 VB6 (`Utility.bas`)**
```vb
Public Sub PUT_DLE_ETX_BCC()
    sTemp = StrConv(AGV_TempStr, vbFromUnicode)
    lLen = LenB(sTemp)
    Send_AGV_Str = StrConv(AGV_TempStr, vbFromUnicode)
    ReDim Preserve Send_AGV_Str(UBound(Send_AGV_Str) + 3)
    Send_AGV_Str(lLen) = DLE
    Send_AGV_Str(lLen + 1) = ETX
    BCC = 0
    For i = 0 To lLen + 1
        BCC = BCC Xor Send_AGV_Str(i)
    Next
    Send_AGV_Str(lLen + 2) = BCC
End Sub
```

**C# 對應**
```csharp
private const byte DLE = 0x10;
private const byte ETX = 0x03;

public static byte[] BuildAgvPacket(string message)
{
    byte[] body = Encoding.Default.GetBytes(message);
    byte[] packet = new byte[body.Length + 3];
    Array.Copy(body, packet, body.Length);
    packet[body.Length]     = DLE;
    packet[body.Length + 1] = ETX;

    byte bcc = 0;
    for (int i = 0; i <= body.Length + 1; i++)
        bcc ^= packet[i];
    packet[body.Length + 2] = bcc;

    return packet;
}
```

---

### 2.8 INI 設定檔讀取

**原始 VB6 (`Utility.bas`)**
```vb
Public Function IniRead(lpAppName As String, lpKeyName As String, lpFileName As String) As String
    valid = GetPrivateProfileString(lpAppName, lpKeyName, lpDefault, lpReturnString, size, lpFileName)
    IniRead = Left$(lpReturnString, valid)
End Function
```

**C# 建議改用 JSON 設定檔**
```csharp
// appsettings.json
{
  "Database": { "ConnectionString": "..." },
  "ComPorts": { "AGV": 9, "PLC_Start": 4, "PLC_End": 8 }
}

// 讀取
var config = new ConfigurationBuilder()
    .AddJsonFile("appsettings.json")
    .Build();
string connStr = config["Database:ConnectionString"];
```

> 若需維持 INI 相容，可繼續 P/Invoke `GetPrivateProfileString` 或使用第三方套件 `IniFile`。

---

### 2.9 全域常數轉換

**原始 VB6 (`Api232-b.bas`, `Global.bas`)**
```vb
Global Const B9600 = &HC
Global Const BIT_8 = &H3
Global Const P_EVEN = &H18
Public Const HOST = "1"
Public Const AGV  = "2"
```

**C# 對應**
```csharp
public static class SerialConstants
{
    public const int B9600   = 0x0C;
    public const int BIT_8   = 0x03;
    public const int P_EVEN  = 0x18;
}

public static class AgvProtocol
{
    public const string HOST  = "1";
    public const string AGV   = "2";
    public const string DUMMY = "0";
    public const string READY = "1";
}
```

---

### 2.10 Timer 事件

**原始 VB6**
```vb
Begin VB.Timer TimerNEW
    Interval = 600
End
Private Sub TimerNEW_Tick()
    ...
End Sub
```

**C# WinForms**
```csharp
private System.Windows.Forms.Timer _timerNew;

_timerNew = new System.Windows.Forms.Timer();
_timerNew.Interval = 600;
_timerNew.Tick += TimerNew_Tick;
_timerNew.Start();

private void TimerNew_Tick(object sender, EventArgs e)
{
    // PLC 輪詢邏輯
}
```

---

## 三、整體轉換策略建議

### 方案比較

| 方案 | 說明 | 優點 | 缺點 |
|------|------|------|------|
| **A. 手動重寫（C# WinForms）** | 逐函式轉換至 C# WinForms | 完整控制、效能最佳 | 工時最高（估 3–6 個月） |
| **B. 自動工具輔助（VBUC）** | 使用 Visual Basic Upgrade Companion 自動轉換 | 速度快，覆蓋率約 70–80% | 需購買授權，仍需大量手工修正 |
| **C. 逐步遷移** | 保留 VB6 UI，將業務邏輯先抽至 C# COM 元件 | 風險低，可分階段 | 架構複雜，長期維護負擔大 |

**建議採用方案 A（手動重寫）**，理由：
- SK900_plc.bas 邏輯龐大（355KB），自動轉換品質難以保證
- 串列埠通訊可改用現代 `SerialPort` API，徹底移除對 api232.dll 的依賴
- C# WinForms 架構清晰，後續可進一步升級至 WPF 或 .NET MAUI

---

## 四、建議轉換順序

```
1. 建立 C# WinForms 解決方案（.NET 8 或 .NET 6 LTS）
2. 轉換常數與型別定義（Global.bas → Constants.cs、Types.cs）
3. 轉換資料庫層（ADODB → SqlClient / EF Core）
4. 轉換 INI 讀取（Utility.bas → appsettings.json）
5. 轉換串列埠封裝（Api232-b.bas → SerialPortWrapper.cs）
6. 轉換 PLC 通訊邏輯（SK900_plc.bas → PlcCommunicator.cs）
7. 轉換 AGV 通訊邏輯（MDImainForm.frm 相關段落 → AgvCommunicator.cs）
8. 重建 UI 表單（MDImainForm, frmComPort, frmMsgBox）
9. 整合測試與驗收
```

---

## 五、注意事項

1. **字元編碼**：原始碼中有大量繁體中文（Big5，Charset=136），轉換時需確認資料庫和通訊協定的編碼一致性（`Encoding.GetEncoding(950)`）。
2. **32-bit vs 64-bit**：api232.dll 為 32-bit DLL，若要繼續使用 P/Invoke 必須以 x86 目標平台編譯；若改用 `SerialPort` 則可編譯為 x64 或 AnyCPU。
3. **`On Error Resume Next` → try/catch**：VB6 的錯誤處理需全面改寫為 C# 的結構化例外處理。
4. **`DoEvents` → async/await**：VB6 中大量使用 `DoEvents` 維持 UI 回應，C# 應改為非同步模式。
5. **固定長度字串陣列**（`String * N`）：在通訊協定相關的程式碼中需特別注意，確保位元組對齊。
