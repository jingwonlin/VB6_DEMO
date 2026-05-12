using System.Data;
using System.Text;
using GF3Com.Constants;
using GF3Com.Database;

namespace GF3Com.Communication;

/// <summary>
/// PLC 通訊邏輯 (from TimerNEW_Timer / dr_sub in MDImainForm.frm)
/// 輪詢 COM4–COM8，解析回應並更新資料庫。
/// </summary>
internal sealed class PlcCommunicator
{
    private readonly SerialPortWrapper[] _ports;  // index 4–8
    private readonly DatabaseContext _db;
    private readonly Action<int, string, string, string, Color> _updateUi;

    // 目前輪詢的 port index (4–8，對應 iPortIndex in VB6)
    public int CurrentPortIndex { get; private set; } = 4;

    /// <summary>
    /// </summary>
    /// <param name="ports">長度 9，有效索引 4–8</param>
    /// <param name="db">已開啟的資料庫連線</param>
    /// <param name="updateUi">更新 frmComPort 標籤的委派 (portIndex+17, len, str, status, color)</param>
    public PlcCommunicator(SerialPortWrapper[] ports,
                           DatabaseContext db,
                           Action<int, string, string, string, Color> updateUi)
    {
        _ports    = ports;
        _db       = db;
        _updateUi = updateUi;
    }

    /// <summary>
    /// 每次 Timer 觸發時執行 (對應 TimerNEW_Timer 主體)
    /// </summary>
    public void Poll()
    {
        int pi = CurrentPortIndex;

        if (pi <= 8)
        {
            // ── 步驟 1：先執行 dr_sub（對應 VB6 TimerNEW_Timer 最前面的 5 個 If 判斷）
            // port 4→awno"1", 5→"2", 6→"3", 7→"4", 8→"5"
            string awno = (pi - 3).ToString();
            DrSub(awno, pi);

            // ── 步驟 2：接著執行一般 PLC 狀態輪詢
            // 原料 port 每次都送查詢
            if (_ports[8].IsOpen)
                _ports[8].Output(PlcConstants.GetPlcQuery(8));

            if (_ports[pi].IsOpen && pi != 8)
                _ports[pi].Output(PlcConstants.GetPlcQuery(pi));
            else if (!_ports[pi].IsOpen)
                _updateUi(pi + 17, "", "", "通訊埠未開啟!!", Color.Red);

            // 收資料
            var (jj, sBuf) = ReadUntilExpected(pi);
            int expectedLen = Globals.PlcLen[pi + 17];

            bool validLen = pi switch
            {
                8 => jj == 50 || jj == 200,
                4 => jj >= 0,
                5 => jj == 20,
                6 => jj == 14,
                7 => jj == 16,
                _ => false
            };

            if (validLen)
            {
                string trimmed = sBuf.Length >= expectedLen
                    ? sBuf[..expectedLen]
                    : sBuf;

                bool headerOk = pi == 4 && jj == 28
                    || trimmed.StartsWith(">5")
                    || trimmed.StartsWith(">2");

                if (headerOk)
                {
                    // 更新特殊站台負載欄位
                    if (pi == 5 && jj == 20)
                    {
                        _db.Execute($"update st set st_load='{SafeChar(sBuf, 15)}' where substring(st_stno,1,4)='B101'");
                        _db.Execute($"update st set st_load='{SafeChar(sBuf, 16)}' where substring(st_stno,1,4)='B102'");
                        _db.Execute($"update st set st_load='{SafeChar(sBuf, 17)}' where substring(st_stno,1,4)='B103'");
                    }
                    if (pi == 6 && jj == 14)
                        _db.Execute($"update st set st_load='{SafeChar(sBuf, 11)}' where substring(st_stno,1,4)='C102'");

                    UpdateSt(pi + 17, trimmed);
                    _updateUi(pi + 17, jj.ToString(), trimmed, "資料正確!!", Color.Blue);
                }
                else
                {
                    _updateUi(pi + 17, jj.ToString(), trimmed, "資料不正確!!", Color.Red);
                    Thread.Sleep(250);
                    _ports[pi].DiscardInBuffer();
                }
            }
            else
            {
                _updateUi(pi + 17, jj.ToString(), sBuf, "資料長度不符!!", Color.Red);
                Thread.Sleep(250);
                _ports[pi].DiscardInBuffer();
            }
        }
        else if (pi == 9)
        {
            // 檢查天車警報燈號 (port index 9)
            CheckCraneAlarms();
        }

        // 輪流切換 port (4→5→6→7→8→4…)
        CurrentPortIndex = CurrentPortIndex >= 8 ? 4 : CurrentPortIndex + 1;
    }

    // ── 私有方法 ────────────────────────────────────────────────────

    /// <summary>輪詢等待直到收到期望長度的資料</summary>
    private (int length, string buffer) ReadUntilExpected(int portIndex)
    {
        var sb  = new StringBuilder(256);
        int jj  = 0, ii = 0;
        int maxLen = portIndex switch { 4=>28, 5=>20, 6=>14, 7=>16, 8=>50, _=>50 };

        while (true)
        {
            byte[] chunk = _ports[portIndex].ReadAvailable();
            if (chunk.Length > 0)
            {
                string s = Encoding.GetEncoding(950).GetString(chunk);
                sb.Append(s);
                jj += chunk.Length;
            }
            else
            {
                ii++;
            }

            bool done = ii > 5
                || (portIndex == 4 && jj >= 28)
                || (portIndex == 5 && jj >= 20)
                || (portIndex == 6 && jj >= 14)
                || (portIndex == 7 && jj >= 16)
                || (portIndex == 8 && jj >= 50);
            if (done) break;

            Thread.Sleep(200);
        }
        return (jj, sb.ToString());
    }

    /// <summary>
    /// 根據 port + 17 index 更新對應站台狀態 (UpdateST in VB6)
    /// </summary>
    private void UpdateSt(int stIndex, string buf)
    {
        (string[] stNo, int[] stByte) = stIndex switch
        {
            21 => (Globals.ST21_StNo, Globals.ST21_StByte),
            22 => (Globals.ST22_StNo, Globals.ST22_StByte),
            23 => (Globals.ST23_StNo, Globals.ST23_StByte),
            24 => (Globals.ST24_StNo, Globals.ST24_StByte),
            25 => (Globals.ST25_StNo, Globals.ST25_StByte),
            _  => (Array.Empty<string>(), Array.Empty<int>())
        };

        int count = stNo.Length - 1;
        for (int i = 1; i <= count; i++)
        {
            if (string.IsNullOrEmpty(stNo[i])) continue;
            if (stByte[i] > buf.Length) continue;

            string loadVal = buf[stByte[i] - 1].ToString();
            _db.Execute(
                $"update st set st_load='{loadVal}' " +
                $"where substring(st_stno,1,4)='{stNo[i]}'");
        }
    }

    /// <summary>天車警報燈號檢查 (from TimerNEW_Timer iPortIndex==9 section)</summary>
    private void CheckCraneAlarms()
    {
        var dt = _db.Query(
            "Select cr_awno,cr_err from cr where cr_awno < '6'");

        foreach (DataRow row in dt.Rows)
        {
            int awno = Convert.ToInt32(row["cr_awno"]);
            int err  = Convert.ToInt32(row["cr_err"]);
            Globals.CraneErrNo[awno] = err;
        }

        for (int i = 1; i <= 4; i++)
        {
            if (Globals.CraneErrNo[i] == Globals.CraneTemp[i]) continue;
            if (!_ports[i + 3].IsOpen) continue;

            string cmd = Globals.CraneErrNo[i] == 0
                ? ">500\r\n"
                : (Globals.CraneErrNo[i] == 67 || Globals.CraneErrNo[i] >= 3000)
                    ? ">510\r\n"
                    : string.Empty;

            if (!string.IsNullOrEmpty(cmd))
            {
                _ports[i + 3].Output(cmd);
                Thread.Sleep(500);
                _ports[i + 3].DiscardInBuffer();
            }
            Globals.CraneTemp[i] = Globals.CraneErrNo[i];
        }
    }

    private static char SafeChar(string s, int oneBasedIndex)
    {
        int idx = oneBasedIndex - 1;
        return idx >= 0 && idx < s.Length ? s[idx] : '0';
    }

    /// <summary>
    /// dr_sub 邏輯：讀取 DR 設備狀態並更新資料庫 (from dr_sub in MDImainForm.frm)
    /// </summary>
    public bool DrSub(string awno, int plcPort)
    {
        _ports[plcPort].DiscardInBuffer();
        string id = "21";

        if (_ports[plcPort].IsOpen)
            _ports[plcPort].Output($"{id}{awno}{awno}\r\n");

        var (jj, drBuf) = ReadWithTimeout(plcPort, 12);

        _updateUi(plcPort + 17, jj.ToString(), drBuf, string.Empty, Color.Blue);

        if (jj != 12)
        {
            string msg = $"ID{id},{jj}:{drBuf[..Math.Min(10, drBuf.Length)]}";
            _updateUi(plcPort + 17, jj.ToString(), drBuf, msg, Color.Red);
            SysLog("dr_sub", msg);
            Thread.Sleep(1000);
            return false;
        }

        // 驗證字元均為數字
        for (int i = 1; i < jj - 2; i++)
        {
            if (drBuf[i] < '0' || drBuf[i] > '9')
            {
                _updateUi(plcPort + 17, jj.ToString(), drBuf,
                    $"awno{awno},ID{id}: 第{i+1}字元錯", Color.Red);
                Thread.Sleep(1000);
                return false;
            }
        }

        // 更新 dr 表
        // VB6: Mid(dr_buf, 1, 1) = drBuf[0], Mid(dr_buf, 2, 9) = drBuf[1..10] (9 chars)
        _db.Execute1(
            $"update dr set dr_use='{drBuf[0]}',dr_stat='{drBuf[1..10]}' " +
            $"where dr_awno='{awno}'");
        _updateUi(plcPort + 17, jj.ToString(), drBuf, $"ID{id}: OK", Color.Blue);

        // 查詢待處理命令
        var dt = _db.Query(
            $"select * from dr where dr_awno='{awno}' " +
            "and (substring(dr_cvar1,1,2)='01' or substring(dr_cvar1,1,2)='11' " +
            "or substring(dr_cvar1,1,2)='21')");

        if (dt.Rows.Count == 0)
        {
            Thread.Sleep(1000);
            return true;
        }

        DataRow brs2 = dt.Rows[0];
        string cvar1 = brs2["dr_cvar1"].ToString() ?? "";
        string prefix = cvar1[..2];

        id = prefix switch
        {
            "01" => "22",
            "11" => "23",
            "21" => "25",
            _    => "22"
        };

        _ports[plcPort].DiscardInBuffer();
        if (_ports[plcPort].IsOpen)
            _ports[plcPort].Output($"{id}{awno}{awno}\r\n");

        var (jj2, libBuf) = ReadWithTimeout(plcPort, 4);
        _updateUi(plcPort + 17, jj2.ToString(), libBuf, string.Empty, Color.Blue);

        if (jj2 != 4)
        {
            string msg = $"ID{id},{jj2}:{drBuf[..Math.Min(10, drBuf.Length)]}";
            _updateUi(plcPort + 17, jj2.ToString(), libBuf, msg, Color.Red);
            SysLog("dr_sub1", msg);
            return false;
        }

        if (libBuf[..2] == "11")
        {
            string newStat = prefix switch
            {
                "01" => "02",
                "11" => "12",
                "21" => "22",
                _    => cvar1
            };
            _db.Execute1(
                $"update dr set dr_cvar1='{newStat}' where dr_awno='{awno}'");
            Thread.Sleep(1000);
        }
        else if (libBuf[..2] == "19")
        {
            _db.Execute1($"update dr set dr_cvar1='19' where dr_awno='{awno}'");
        }
        else
        {
            SysLog("dr_sub2", $"ID{id},{jj2}:{drBuf[..Math.Min(10, drBuf.Length)]}");
            _updateUi(plcPort + 17, jj2.ToString(), libBuf,
                $"ID{id},{jj2}:{libBuf[..Math.Min(10, libBuf.Length)]}", Color.Red);
            Thread.Sleep(1000);
        }

        return true;
    }

    private (int length, string buffer) ReadWithTimeout(int portIndex, int expectLen)
    {
        var sb = new StringBuilder(256);
        int jj = 0, ii = 0;

        while (true)
        {
            byte[] chunk = _ports[portIndex].ReadAvailable();
            if (chunk.Length > 0)
            {
                sb.Append(Encoding.GetEncoding(950).GetString(chunk));
                jj += chunk.Length;
            }
            else
            {
                ii++;
            }
            if (ii > 5 || jj >= expectLen) break;
            Thread.Sleep(100);
        }
        return (jj, sb.ToString());
    }

    private static void SysLog(string tag, string msg)
    {
        // 對應 VB6 的 syslog 呼叫，寫入 LogMessage 目錄
        string logDir  = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "LogMessage");
        string logFile = Path.Combine(logDir, $"{DateTime.Today:yyyyMMdd}.log");
        Directory.CreateDirectory(logDir);
        File.AppendAllText(logFile, $"[{DateTime.Now:HH:mm:ss}][{tag}] {msg}{Environment.NewLine}");
    }
}
