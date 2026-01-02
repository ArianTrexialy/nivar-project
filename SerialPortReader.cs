using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DeviceAnalisys_v5
{


    public class SerialPortReader
    {
        private SerialPort _port;
        private List<byte> _buffer = new List<byte>();
        private int _timeIndex = 0;
        public bool IsConnected => _port != null && _port.IsOpen;

        public SerialPortReader(string portName)
        {
            _port = new SerialPort(portName, 115200, Parity.None, 8, StopBits.One);
            _port.ReadBufferSize = 1_000_000;
            _port.ReadTimeout = 500;

            _port.DataReceived += Port_DataReceived;
        }
        public void Start()
        {
            if (_port != null && _port.IsOpen) { _port.Close(); }
            _port.Open();
        }

        public void Stop()
        {
            if (_port.IsOpen) _port.Close();
            _buffer.Clear();
            _timeIndex = 0; 
        }
        private void Port_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            int bytes = _port.BytesToRead;
            byte[] temp = new byte[bytes];
            _port.Read(temp, 0, bytes);

            _buffer.AddRange(temp);

            ExtractData();
        }

        private void ExtractData()
        {
            while (_buffer.Count >= 32)
            {
                for (int i = 0; i < 4; i++)
                {
                    int offset = i * 8;
                    ushort USetPoint = (ushort)((_buffer[offset + 0]) | _buffer[offset + 1] << 8);
                    ushort UActual = (ushort)((_buffer[offset + 2]) | _buffer[offset + 3] << 8);
                    ushort Uval1 = (ushort)((_buffer[offset + 4]) | _buffer[offset + 5] << 8);
                    ushort Uval2 = (ushort)((_buffer[offset + 6]) | _buffer[offset + 7] << 8);
                    DeviceData data = new DeviceData
                    {
                        TestID = i + 1,
                        Time = _timeIndex,
                        SetPoint = NormalizeValue(USetPoint),
                        Actual = NormalizeValue(UActual),
                        Pitch = NormalizeValue(Uval1),
                        Roll = NormalizeValue(Uval2)
                    };
                    GlobalData.DiagramQueue.Enqueue(data);
                    // حذف این خط: GlobalData.DBList.Add(data);  // چون duplicate میشه و سریال ست نشده
                }
                _timeIndex++;
                _buffer.RemoveRange(0, 32);
            }
        }
        double NormalizeValue(ushort value)
        {
            double val = value;
            if (val > 32767) val = 32768 - val;
            return val / 1000.0;
        }
        public void SendCommand(string command)
        {
            try
            {
                if (_port != null && _port.IsOpen)
                {
                    _port.Write(command);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Send command failed: {ex.Message}");
            }
        }
    }
}
