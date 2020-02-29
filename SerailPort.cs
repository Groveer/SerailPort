using System;
using System.Collections;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace GR
{
  /*
  使用例子：
      SerialPort comm=new SerialPort();
      comm.ShowErrMsg+=(s,m)=>MessageBox.Show(m);
      Hashtable Ht = new Hashtable();
      Ht.Add(SerialPort.KEY_PORT, 2);
      //初始化串口
      comm.InitComm(Ht);
      comm.Recived += Comm_Recived;
  */
  public class SerialPort
  {
    [DllImport("PComm.dll", EntryPoint = "sio_open")]
    private static extern int Sio_open(int port);

    [DllImport("PComm.dll", EntryPoint = "sio_ioctl")]
    private static extern int Sio_ioctl(int port, int baud, int mode);

    [DllImport("PComm.dll", EntryPoint = "sio_DTR")]
    private static extern int Sio_DTR(int port, int mode);

    [DllImport("PComm.dll", EntryPoint = "sio_RTS")]
    private static extern int Sio_RTS(int port, int mode);

    [DllImport("PComm.dll", EntryPoint = "sio_close")]
    private static extern int Sio_close(int port);

    [DllImport("PComm.dll", EntryPoint = "sio_read")]
    private static extern int Sio_read(int port, ref byte buf, int length);

    [DllImport("PComm.dll", EntryPoint = "sio_write")]
    private static extern int Sio_write(int port, ref byte buf, int length);

    [DllImport("PComm.dll", EntryPoint = "sio_SetReadTimeouts")]
    private static extern int Sio_SetReadTimeouts(int port, int TotalTimeouts, int IntervalTimeouts);

    [DllImport("PComm.dll", EntryPoint = "sio_SetWriteTimeouts")]
    private static extern int Sio_SetWriteTimeouts(int port, int TotalTimeouts, int IntervalTimeouts);

    [DllImport("PComm.dll", EntryPoint = "sio_AbortRead")]
    private static extern int Sio_AbortRead(int port);

    [DllImport("PComm.dll", EntryPoint = "sio_AbortWrite")]
    private static extern int Sio_AbortWrite(int port);

    [DllImport("PComm.dll", EntryPoint = "sio_getmode")]
    private static extern int Sio_getmode(int port);

    [DllImport("PComm.dll", EntryPoint = "sio_getbaud")]
    private static extern int Sio_getbaud(int port);

    [DllImport("PComm.dll", EntryPoint = "sio_flowctrl")]
    private static extern int Sio_flowctrl(int port, int mode);

    [DllImport("PComm.dll", EntryPoint = "sio_flush")]
    private static extern int Sio_flush(int port, int mode);

    // Function return error code
    private const int SIO_OK = 0;
    private const int SIO_BADPORT = -1;
    private const int SIO_OUTCONTROL = -2;
    private const int SIO_NODATA = -4;
    private const int SIO_OPENFAIL = -5;
    private const int SIO_RTS_BY_HW = -6;
    private const int SIO_BADPARAM = -7;
    private const int SIO_WIN32FAIL = -8;
    private const int SIO_BOARDNOTSUPPORT = -9;
    private const int SIO_ABORT_WRITE = -11;
    private const int SIO_WRITETIMEOUT = -12;

    // Self Define function return error code
    private const int ERR_NOANSWER = -101;

    // Baud rate
    public static int B50 = 0x0;
    public static int B75 = 0x1;
    public static int B110 = 0x2;
    public static int B134 = 0x3;
    public static int B150 = 0x4;
    public static int B300 = 0x5;
    public static int B600 = 0x6;
    public static int B1200 = 0x7;
    public static int B1800 = 0x8;
    public static int B2400 = 0x9;
    public static int B4800 = 0xA;
    public static int B7200 = 0xB;
    public static int B9600 = 0xC;
    public static int B19200 = 0xD;
    public static int B38400 = 0xE;
    public static int B57600 = 0xF;
    public static int B115200 = 0x10;
    public static int B230400 = 0x11;
    public static int B460800 = 0x12;
    public static int B921600 = 0x13;

    // Mode setting Data bits define
    public static int BIT_5 = 0x0;
    public static int BIT_6 = 0x1;
    public static int BIT_7 = 0x2;
    public static int BIT_8 = 0x3;

    // Mode setting Stop bits define
    public static int STOP_1 = 0x0;
    public static int STOP_2 = 0x4;

    // Mode setting Parity define
    public static int P_EVEN = 0x18;
    public static int P_ODD = 0x8;
    public static int P_SPC = 0x38;
    public static int P_MRK = 0x28;
    public static int P_NONE = 0x0;

    // Key name
    public static string KEY_PORT = "Port";
    public static string KEY_BAUDRATE = "Baud_Rate";
    public static string KEY_PARITY = "Parity";
    public static string KEY_BYTESIZE = "Byte_Size";
    public static string KEY_STOPBITS = "Stop_Bits";
    public static string KEY_BEFOREDELAY = "Before_Delay";
    public static string KEY_BYTEDELAY = "Byte_Delay";
    public static string KEY_READINTERVALTIMEOUT = "Read_Interval_Timeout";
    public static string KEY_AFTERDELAY = "After_Delay";

    // Port param
    private int Gl_Int_Port = 1;
    private int Gl_Int_Baudrate = B9600;
    private int Gl_Int_Parity = P_NONE;
    private int Gl_Int_ByteSize = BIT_8;
    private int Gl_Int_StopBits = STOP_1;

    // Delay param
    private int Gl_Int_BeforeDelay = 0;
    private int Gl_Int_ByteDelay = 0;
    private int Gl_Int_ReadIntervalTimeout = 50;
    //private int Gl_Int_WriteIntervalTimeout = 50;
    private int Gl_Int_AfterDealy = 1;

    [DllImport("kernel32")]
    private static extern IntPtr LoadLibrary([MarshalAs(UnmanagedType.LPWStr)] string fileName);

    public event Action<SerialPort, string> ShowErrMsg;
    public SerialPort()
    {
      var dllFile = System.IO.Path.Combine(Environment.Is64BitProcess ? "Win64" : "Win86", "PCOMM.dll");
      dllFile = System.IO.Directory.GetCurrentDirectory() + "\\" + dllFile;
      LoadLibrary(dllFile);
    }
    /// <summary>
    /// 解析通讯参数
    /// </summary>
    /// <param name="Hb_CommParam"></param>
    private void AnalyseCommParam(Hashtable Ht_CommParam)
    {
      //Port
      if (Ht_CommParam.Contains(KEY_PORT))
      {
        Gl_Int_Port = Convert.ToInt32(Ht_CommParam[KEY_PORT]);
      }
      //Baud rate
      if (Ht_CommParam.Contains(KEY_BAUDRATE))
      {
        Gl_Int_Baudrate = Convert.ToInt32(Ht_CommParam[KEY_BAUDRATE]);
      }
      //Parity
      if (Ht_CommParam.Contains(KEY_PARITY))
      {
        Gl_Int_Parity = Convert.ToInt32(Ht_CommParam[KEY_PARITY]);
      }
      //Byte Size
      if (Ht_CommParam.Contains(KEY_BYTESIZE))
      {
        Gl_Int_ByteSize = Convert.ToInt32(Ht_CommParam[KEY_BYTESIZE]);
      }
      //Stop Bits
      if (Ht_CommParam.Contains(KEY_STOPBITS))
      {
        Gl_Int_StopBits = Convert.ToInt32(Ht_CommParam[KEY_STOPBITS]);
      }
      //Before Delay
      if (Ht_CommParam.Contains(KEY_BEFOREDELAY))
      {
        int.TryParse(Ht_CommParam[KEY_BEFOREDELAY].ToString(), out Gl_Int_BeforeDelay);
      }
      //Byte Delay
      if (Ht_CommParam.Contains(KEY_BYTEDELAY))
      {
        int.TryParse(Ht_CommParam[KEY_BYTEDELAY].ToString(), out Gl_Int_ByteDelay);
      }
      //Read Interval Timeout
      if (Ht_CommParam.Contains(KEY_READINTERVALTIMEOUT))
      {
        int.TryParse(Ht_CommParam[KEY_READINTERVALTIMEOUT].ToString(), out Gl_Int_ReadIntervalTimeout);
      }
      //After Delay
      if (Ht_CommParam.Contains(KEY_AFTERDELAY))
      {
        int.TryParse(Ht_CommParam[KEY_AFTERDELAY].ToString(), out Gl_Int_AfterDealy);
      }
    }
    /// <summary>
    /// 初始化串口通讯
    /// </summary>
    /// <param name="Hb_CommParam"></param>
    public void InitComm(Hashtable Ht_CommParam)
    {
      AnalyseCommParam(Ht_CommParam);

      //Open port
      int i_RtnCode = Sio_open(Gl_Int_Port);
      if (i_RtnCode != SIO_OK)
      {
        ShowErrMsg?.Invoke(this, GetCommErrMsg(i_RtnCode));
        return;
      }

      //Configure communication parameters
      int mode = Gl_Int_Parity | Gl_Int_ByteSize | Gl_Int_StopBits;
      i_RtnCode = Sio_ioctl(Gl_Int_Port, Gl_Int_Baudrate, mode);
      if (i_RtnCode != SIO_OK)
      {
        ShowErrMsg?.Invoke(this, GetCommErrMsg(i_RtnCode));
        return;
      }

      //Flow control
      i_RtnCode = Sio_flowctrl(Gl_Int_Port, 0);
      if (i_RtnCode != SIO_OK)
      {
        ShowErrMsg?.Invoke(this, GetCommErrMsg(i_RtnCode));
        return;
      }

      //DTR
      i_RtnCode = Sio_DTR(Gl_Int_Port, 1);
      if (i_RtnCode != SIO_OK)
      {
        ShowErrMsg?.Invoke(this, GetCommErrMsg(i_RtnCode));
        return;
      }

      //RTS
      i_RtnCode = Sio_RTS(Gl_Int_Port, 1);
      if (i_RtnCode != SIO_OK)
      {
        ShowErrMsg?.Invoke(this, GetCommErrMsg(i_RtnCode));
        return;
      }

      //Set timeout values for sio_read 
      Sio_SetReadTimeouts(Gl_Int_Port, Gl_Int_AfterDealy, Gl_Int_ReadIntervalTimeout);
      //Sio_SetWriteTimeouts(Gl_Int_Port, Gl_Int_BeforeDelay, Gl_Int_WriteIntervalTimeout);
      ReadyRead();
    }

    /// <summary>
    /// 发送十六进值字符串帧，
    /// </summary>
    /// <param name="Str_Data">发送帧</param>
    public async void SendData(string Str_Data)
    {
      int i_RtnCode;
      Str_Data = Str_Data.Replace(" ", "");
      Str_Data = Str_Data.Replace("-", "");
      Str_Data += Str_Data.Length % 2 != 0 ? "0" : "";
      int i_Length = Str_Data.Length / 2;
      byte[] buffer = new byte[i_Length];
      for (int i = 0; i <= i_Length - 1; i++)
      {
        buffer[i] = Convert.ToByte(Str_Data.Substring(i * 2, 2), 16);
      }

      //Write data
      if (Gl_Int_ByteDelay == 0)
      {
        i_RtnCode = Sio_write(Gl_Int_Port, ref buffer[0], i_Length);
        if (i_RtnCode < 0)
        {
          ShowErrMsg?.Invoke(this, GetCommErrMsg(i_RtnCode));
          return;
        }
      }
      else
      {
        for (int i = 0; i <= i_Length - 1; i++)
        {
          i_RtnCode = Sio_write(Gl_Int_Port, ref buffer[i], 1);
          if (i_RtnCode < 0)
          {
            ShowErrMsg?.Invoke(this, GetCommErrMsg(i_RtnCode));
            return;
          }
          await Task.Delay(Gl_Int_ByteDelay);
        }
      }
    }
    public async void SendData(byte[] Data)
    {
      int i_RtnCode;
      int i_Length = Data.Length;
      //Write data
      if (Gl_Int_ByteDelay == 0)
      {
        i_RtnCode = Sio_write(Gl_Int_Port, ref Data[0], i_Length);
        if (i_RtnCode < 0)
        {
          ShowErrMsg?.Invoke(this, GetCommErrMsg(i_RtnCode));
          return;
        }
      }
      else
      {
        for (int i = 0; i <= i_Length - 1; i++)
        {
          i_RtnCode = Sio_write(Gl_Int_Port, ref Data[i], 1);
          if (i_RtnCode < 0)
          {
            ShowErrMsg?.Invoke(this, GetCommErrMsg(i_RtnCode));
            return;
          }
          await Task.Delay(Gl_Int_ByteDelay);
        }
      }
    }
    public event Action<SerialPort, byte[]> Recived;
    private void ReadyRead()
    {
      Task.Run(() => {
        byte[] recbuffer = new byte[4096];
        while (true)
        {
          int i_RtnCode = Sio_read(Gl_Int_Port, ref recbuffer[0], 4096);
          if (i_RtnCode > 0)
          {
            byte[] data = new byte[i_RtnCode];
            for (int i = 0; i < i_RtnCode; i++)
            {
              data[i] = recbuffer[i];
            }
            Recived?.Invoke(this, data);
          }
          Sio_flush(Gl_Int_Port, 1);
        }
      });
    }
    /// <summary>
    /// 关闭串口通讯
    /// </summary>
    public void CloseComm()
    {
      int i_RtnCode = Sio_close(Gl_Int_Port);
      if (i_RtnCode != SIO_OK)
      {
        ShowErrMsg?.Invoke(this, GetCommErrMsg(i_RtnCode));
        return;
      }
    }
    ~SerialPort()
    {
      CloseComm();
    }
    /// <summary>
    /// 获取通讯错误消息
    /// </summary>
    /// <param name="i_ErrCode">错误码</param>
    /// <returns>错误消息</returns>
    private string GetCommErrMsg(int i_ErrCode)
    {
      switch (i_ErrCode)
      {
        case SIO_OK: return "成功";
        case SIO_BADPORT: return "串口号无效,检测串口号!";
        case SIO_OUTCONTROL: return "主板不是MOXA兼容的智能主板!";
        case SIO_NODATA: return "没有可读的数据!";
        case SIO_OPENFAIL: return "打开串口失败,检查串口是否被占用!";
        case SIO_RTS_BY_HW: return "不能控制串口因为已经通过sio_flowctrl设定为自动H/W流控制";
        case SIO_BADPARAM: return "串口参数错误,检查串口参数!";
        case SIO_WIN32FAIL: return "调用Win32函数失败!";
        case SIO_BOARDNOTSUPPORT: return "串口不支持这个函数!";
        case SIO_ABORT_WRITE: return "用户终止写数据块!";
        case SIO_WRITETIMEOUT: return "写数据超时!";
        case ERR_NOANSWER: return "无应答!";
        default: return i_ErrCode.ToString();
      }
    }
  }
  public class CRC
  {

    #region  CRC16
    public static byte[] CRC16(byte[] data)
    {
      int len = data.Length;
      if (len > 0)
      {
        ushort crc = 0xFFFF;

        for (int i = 0; i < len; i++)
        {
          crc = (ushort)(crc ^ (data[i]));
          for (int j = 0; j < 8; j++)
          {
            crc = (crc & 1) != 0 ? (ushort)((crc >> 1) ^ 0xA001) : (ushort)(crc >> 1);
          }
        }
        byte hi = (byte)((crc & 0xFF00) >> 8);  //高位置
        byte lo = (byte)(crc & 0x00FF);         //低位置

        return new byte[] { hi, lo };
      }
      return new byte[] { 0, 0 };
    }
    #endregion

    #region  ToCRC16
    public static string ToCRC16(string content)
    {
      return ToCRC16(content, Encoding.UTF8);
    }

    public static string ToCRC16(string content, bool isReverse)
    {
      return ToCRC16(content, Encoding.UTF8, isReverse);
    }

    public static string ToCRC16(string content, Encoding encoding)
    {
      return ByteToString(CRC16(encoding.GetBytes(content)), true);
    }

    public static string ToCRC16(string content, Encoding encoding, bool isReverse)
    {
      return ByteToString(CRC16(encoding.GetBytes(content)), isReverse);
    }

    public static string ToCRC16(byte[] data)
    {
      return ByteToString(CRC16(data), true);
    }

    public static string ToCRC16(byte[] data, bool isReverse)
    {
      return ByteToString(CRC16(data), isReverse);
    }
    #endregion

    #region  ToModbusCRC16
    public static string ToModbusCRC16(string s)
    {
      return ToModbusCRC16(s, true);
    }

    public static string ToModbusCRC16(string s, bool isReverse)
    {
      return ByteToString(CRC16(StringToHexByte(s)), isReverse);
    }

    public static string ToModbusCRC16(byte[] data)
    {
      return ToModbusCRC16(data, true);
    }

    public static string ToModbusCRC16(byte[] data, bool isReverse)
    {
      return ByteToString(CRC16(data), isReverse);
    }
    #endregion

    #region  ByteToString
    public static string ByteToString(byte[] arr, bool isReverse)
    {
      try
      {
        byte hi = arr[0], lo = arr[1];
        return Convert.ToString(isReverse ? hi + lo * 0x100 : hi * 0x100 + lo, 16).ToUpper().PadLeft(4, '0');
      }
      catch (Exception ex) { throw (ex); }
    }

    public static string ByteToString(byte[] arr)
    {
      try
      {
        return ByteToString(arr, true);
      }
      catch (Exception ex) { throw (ex); }
    }
    //public static string BytesToString(byte[] arr,bool isReverse)
    //{
    //    BitConverter.ToString(arr);
    //}
    #endregion

    #region  StringToHexString
    public static string StringToHexString(string str)
    {
      StringBuilder s = new StringBuilder();
      foreach (short c in str.ToCharArray())
      {
        s.Append(c.ToString("X4"));
      }
      return s.ToString();
    }
    #endregion

    #region  StringToHexByte
    private static string ConvertChinese(string str)
    {
      StringBuilder s = new StringBuilder();
      foreach (short c in str.ToCharArray())
      {
        if (c <= 0 || c >= 127)
        {
          s.Append(c.ToString("X4"));
        }
        else
        {
          s.Append((char)c);
        }
      }
      return s.ToString();
    }

    private static string FilterChinese(string str)
    {
      StringBuilder s = new StringBuilder();
      foreach (short c in str.ToCharArray())
      {
        if (c > 0 && c < 127)
        {
          s.Append((char)c);
        }
      }
      return s.ToString();
    }

    /// <summary>
    /// 字符串转16进制字符数组
    /// </summary>
    /// <param name="hex"></param>
    /// <returns></returns>
    public static byte[] StringToHexByte(string str)
    {
      return StringToHexByte(str, false);
    }

    /// <summary>
    /// 字符串转16进制字符数组
    /// </summary>
    /// <param name="str"></param>
    /// <param name="isFilterChinese">是否过滤掉中文字符</param>
    /// <returns></returns>
    public static byte[] StringToHexByte(string str, bool isFilterChinese)
    {
      string hex = isFilterChinese ? FilterChinese(str) : ConvertChinese(str);

      //清除所有空格
      hex = hex.Replace(" ", "");
      hex = hex.Replace("-", "");
      //若字符个数为奇数，补一个0
      hex += hex.Length % 2 != 0 ? "0" : "";

      byte[] result = new byte[hex.Length / 2];
      for (int i = 0, c = result.Length; i < c; i++)
      {
        result[i] = Convert.ToByte(hex.Substring(i * 2, 2), 16);
      }
      return result;
    }
    #endregion

  }
}
