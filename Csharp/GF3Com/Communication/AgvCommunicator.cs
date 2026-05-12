using System.Text;
using GF3Com.Constants;

namespace GF3Com.Communication;

/// <summary>
/// AGV 通訊邏輯 (from PUT_DLE_ETX_BCC / HOST_To_AGV in Utility.bas + MDImainForm.frm)
/// </summary>
internal sealed class AgvCommunicator
{
    private readonly SerialPortWrapper _agvPort;

    public AgvCommunicator(SerialPortWrapper agvPort)
    {
        _agvPort = agvPort;
    }

    /// <summary>
    /// 開啟 AGV 串列埠並送出握手命令 (對應 MDIForm_Load 中的 MSCommAGV 初始化)
    /// </summary>
    public bool Open(int portNumber = 9, int baud = 9600)
    {
        try
        {
            _agvPort.Open(baud, System.IO.Ports.Parity.Even, 8, System.IO.Ports.StopBits.One);
            SendHandshake("K");
            Globals.AgvOpenOk = 1;
            return true;
        }
        catch
        {
            Globals.AgvOpenOk = -1;
            return false;
        }
    }

    /// <summary>
    /// 組裝並送出 AGV 命令封包
    /// 格式: STX + [DLE+] payload + DLE + ETX + BCC
    /// (對應 PUT_DLE_ETX_BCC + sio_putch(AGV_Port, STX) + sio_write)
    /// </summary>
    public void SendCommand(string payload)
    {
        Globals.AgvTempStr = payload;
        byte[] packet = BuildPacket(payload);
        Globals.SendAgvStr = packet;

        // 先送 STX
        _agvPort.Write(new[] { AgvProtocol.STX });
        Thread.Sleep(500);
        // 再送主體
        _agvPort.Write(packet);
    }

    /// <summary>
    /// 送出移動命令 (對應 Command1_Click in frmComPort.frm)
    /// </summary>
    public void SendMoveCommand(string stNoNow, string stNoNext,
                                 string pallet, string dummy = AgvProtocol.DUMMY)
    {
        string payload = AgvProtocol.AGV
                       + AgvProtocol.HOST
                       + "T"
                       + dummy
                       + "0"
                       + pallet.PadLeft(5, '0')
                       + stNoNow.PadRight(5)
                       + stNoNext.PadRight(5)
                       + "1";
        SendCommand(payload);
    }

    /// <summary>
    /// 送出握手 / 初始化命令
    /// </summary>
    public void SendHandshake(string cmd)
    {
        _agvPort.Output(cmd);
    }

    // ── 協定組裝 ────────────────────────────────────────────────────

    /// <summary>
    /// 組裝 DLE+ETX+BCC 封包 (PUT_DLE_ETX_BCC in Utility.bas)
    /// </summary>
    public static byte[] BuildPacket(string message)
    {
        byte[] body   = Encoding.GetEncoding(950).GetBytes(message);
        byte[] packet = new byte[body.Length + 3];
        Array.Copy(body, packet, body.Length);

        packet[body.Length]     = AgvProtocol.DLE;
        packet[body.Length + 1] = AgvProtocol.ETX;

        byte bcc = 0;
        for (int i = 0; i <= body.Length + 1; i++)
            bcc ^= packet[i];
        packet[body.Length + 2] = bcc;

        Globals.Bcc = bcc;
        return packet;
    }

    /// <summary>
    /// 解析 AGV 收到的回應 (AGV_To_HOST in VB6)
    /// </summary>
    public AgvResponse ParseResponse(byte[] data)
    {
        if (data.Length < 3) return new AgvResponse(false, string.Empty);

        // 驗證 BCC
        byte bcc = 0;
        for (int i = 0; i < data.Length - 1; i++)
            bcc ^= data[i];

        bool bccOk = bcc == data[^1];
        string body = Encoding.GetEncoding(950).GetString(data, 0, data.Length - 1);
        return new AgvResponse(bccOk, body);
    }
}

internal record AgvResponse(bool BccOk, string Body);
