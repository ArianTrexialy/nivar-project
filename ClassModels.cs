using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DeviceAnalisys_v5
{
    //class used: udp,db,dataqueue - fill udp to dataqueue
    public class DeviceData
    {
        public int TestID { get; set; }
        public int Time { get; set; }
        public double SetPoint { get; set; }
        public double Actual { get; set; }
        public double Pitch { get; set; }
        public double Roll { get; set; }

        public string SerialNumber { get; set; } = "";
    }

    //
    public static class GlobalData
    {
        public static ConcurrentQueue<DeviceData> DiagramQueue { get; } = new ConcurrentQueue<DeviceData>();

        public static List<DeviceData> DBList { get;} = new List<DeviceData>();
    }
}
