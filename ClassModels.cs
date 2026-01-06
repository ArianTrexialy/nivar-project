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

        //For Raw Data
        public double SetPointRaw { get; set; }
        public double ActualRaw { get; set; }
        public double PitchRaw { get; set; }
        public double RollRaw { get; set; }

        //Scale data /1000
        public double SetPointDeg { get; set; }
        public double ActualDeg { get; set; }
        public double PitchDeg { get; set; }
        public double RollDeg { get; set; }

        public string SerialNumber { get; set; } = "";
    }

    public class StepResult
    {
        public double From { get; set; }
        public double To { get; set; }
        public double DeadTime { get; set; }        // ms
        public double RiseTime { get; set; }         // ms
        public double SettlingTime { get; set; }     // ms
        public double Overshoot { get; set; }        // %
        public double SSE { get; set; }
        public bool IsSensitivity { get; set; }
        public double ActualValue { get; set; }
    }

    //
    public static class GlobalData
    {
        public static double Scale = 1000.0;
        public static List<DeviceData> DBList = new List<DeviceData>();
    }
}
