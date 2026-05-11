using System.IO.Ports;
using System.Text;

namespace GF3Com.Communication;

/// <summary>
/// System.IO.Ports.SerialPort 封裝，取代 MSComm 控制項 (from MDImainForm.frm)
/// 每個 COM port 建立一個實例。
/// </summary>
internal sealed class SerialPortWrapper : IDisposable
{
    private SerialPort? _port;
    private bool _disposed;

    public int PortNumber   { get; }
    public bool IsOpen      => _port?.IsOpen ?? false;
    public int BytesToRead  => _port?.BytesToRead ?? 0;

    /// <summary>收到資料事件 (取代 MSComm.OnComm / RThreshold)</summary>
    public event EventHandler<byte[]>? DataReceived;

    public SerialPortWrapper(int portNumber)
    {
        PortNumber = portNumber;
    }

    /// <summary>
    /// 開啟 COM port (對應 MSCommAR(n).PortOpen = True)
    /// </summary>
    public void Open(int baud = 9600,
                     Parity parity = Parity.Even,
                     int dataBits = 8,
                     StopBits stopBits = StopBits.One,
                     int receivedBytesThreshold = 1)
    {
        _port = new SerialPort($"COM{PortNumber}", baud, parity, dataBits, stopBits)
        {
            DtrEnable = true,
            ReceivedBytesThreshold = receivedBytesThreshold,
            ReadTimeout  = 500,
            WriteTimeout = 500,
            Encoding = Encoding.GetEncoding(950)   // Big5，與 VB6 一致
        };
        _port.DataReceived += OnPortDataReceived;
        _port.Open();
    }

    private void OnPortDataReceived(object sender, SerialDataReceivedEventArgs e)
    {
        if (_port is null || !_port.IsOpen) return;
        int count = _port.BytesToRead;
        if (count <= 0) return;
        byte[] buf = new byte[count];
        _port.Read(buf, 0, count);
        DataReceived?.Invoke(this, buf);
    }

    /// <summary>以 Big5 字串輸出 (對應 MSCommAR(n).Output = sSend)</summary>
    public void Output(string text)
    {
        if (_port is null || !_port.IsOpen) return;
        byte[] bytes = Encoding.GetEncoding(950).GetBytes(text);
        _port.Write(bytes, 0, bytes.Length);
    }

    /// <summary>輸出原始位元組陣列</summary>
    public void Write(byte[] data) =>
        _port?.Write(data, 0, data.Length);

    /// <summary>
    /// 同步讀取（輪詢模式，對應 VB6 的 Do/Loop + InBufferCount/Input）
    /// </summary>
    public byte[] ReadAvailable()
    {
        if (_port is null || !_port.IsOpen || _port.BytesToRead == 0)
            return Array.Empty<byte>();
        int count = _port.BytesToRead;
        byte[] buf = new byte[count];
        _port.Read(buf, 0, count);
        return buf;
    }

    /// <summary>清除輸入緩衝區 (對應 sio_flush / MSComm Input 清除)</summary>
    public void DiscardInBuffer() => _port?.DiscardInBuffer();

    /// <summary>關閉 COM port</summary>
    public void Close()
    {
        if (_port?.IsOpen == true)
            _port.Close();
    }

    public void Dispose()
    {
        if (_disposed) return;
        _port?.Dispose();
        _disposed = true;
    }
}
