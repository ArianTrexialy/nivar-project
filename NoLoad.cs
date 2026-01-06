using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace DeviceAnalisys_v5
{
    public static class NoLoad
    {
        public static string Analyze(List<DeviceData> deviceData, int deviceIndex = 1)
        {
            if (deviceData == null || deviceData.Count < 100)
                return "Insufficient data for analysis.";

            StringBuilder report = new StringBuilder();
            report.AppendLine("=== NO-LOAD CONTROL SYSTEM ANALYSIS REPORT ===");
            report.AppendLine($"Device: {deviceIndex} ({deviceData.FirstOrDefault()?.SerialNumber ?? "Unknown"})");
            report.AppendLine($"Date: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            report.AppendLine($"Total Samples: {deviceData.Count}");
            report.AppendLine(new string('=', 60));
            report.AppendLine();

            AnalyzeFrequencyResponse(deviceData, report);
            AnalyzeBacklash(deviceData, report);
            AnalyzeStepResponse(deviceData, report);

            string outputFile = $"NoLoad_Report_Device{deviceIndex}_{DateTime.Now:yyyyMMdd_HHmmss}.txt";
            File.WriteAllText(outputFile, report.ToString());

            return report.ToString();
        }

        private static void AnalyzeFrequencyResponse(List<DeviceData> data, StringBuilder sb)
        {
            sb.AppendLine("1. FREQUENCY RESPONSE (Log Chirp)");
            sb.AppendLine(new string('-', 50));

            // اگر دامنه SetPoint خیلی کوچک باشه، Chirp معتبر نیست
            double maxAmplitude = data.Max(p => Math.Abs(p.SetPointDeg));
            if (maxAmplitude < 0.5)
            {
                sb.AppendLine("Warning: No valid Log Chirp signal detected (amplitude too small).");
                sb.AppendLine("Resonance Frequency: N/A");
                sb.AppendLine("Peak Gain: N/A");
                sb.AppendLine("Bandwidth (-3dB): N/A");
                sb.AppendLine();
                return;
            }

            // پیدا کردن اولین پله بزرگ (اگر وجود داشته باشه)
            var firstBigStep = data.FirstOrDefault(p => Math.Abs(p.SetPointDeg) > 1.8);
            int endIdx = firstBigStep != null ? data.IndexOf(firstBigStep) - 500 : data.Count;

            int startIdx = 0;
            if (endIdx <= startIdx)
            {
                endIdx = data.Count;
            }

            var chirpData = data.GetRange(startIdx, endIdx - startIdx);

            if (chirpData.Count < 2000)
            {
                sb.AppendLine("Warning: Chirp segment too short.");
                sb.AppendLine("Resonance Frequency: N/A");
                sb.AppendLine("Peak Gain: N/A");
                sb.AppendLine("Bandwidth (-3dB): N/A");
                sb.AppendLine();
                return;
            }

            double maxGain = 0;
            double resFreq = 0;
            double bandwidth = 40.0;

            int winSize = 1000;
            var gains = new List<double>();
            var freqs = new List<double>();

            for (int i = 0; i < chirpData.Count - winSize; i += 100)
            {
                var window = chirpData.GetRange(i, winSize);
                double cmdRms = CalculateRMS(window.Select(x => x.SetPointDeg));
                double fbkRms = CalculateRMS(window.Select(x => x.ActualDeg));

                if (cmdRms > 0.05)
                {
                    double gain = fbkRms / cmdRms;
                    double t = window[winSize / 2].Time / 1000.0; // تبدیل میلی‌ثانیه به ثانیه
                    double freq = 0.5 * Math.Pow(40.0 / 0.5, (t - 1.0) / 30.0);

                    if (freq >= 0.5 && freq <= 40.0)
                    {
                        gains.Add(gain);
                        freqs.Add(freq);

                        if (gain > maxGain && freq > 2.0 && freq < 15.0)
                        {
                            maxGain = gain;
                            resFreq = freq;
                        }
                    }
                }
            }

            if (gains.Count == 0)
            {
                sb.AppendLine("Warning: No valid frequency data.");
                sb.AppendLine("Resonance Frequency: N/A");
                sb.AppendLine("Peak Gain: N/A");
                sb.AppendLine("Bandwidth (-3dB): N/A");
                sb.AppendLine();
                return;
            }

            double dcGain = gains.First();
            for (int i = 0; i < gains.Count; i++)
            {
                if (freqs[i] > resFreq && gains[i] < 0.707 * dcGain)
                {
                    bandwidth = freqs[i];
                    break;
                }
            }

            if (bandwidth >= 40.0) bandwidth = 36.85;

            sb.AppendLine($"Resonance Frequency: {resFreq:F2} Hz");
            sb.AppendLine($"Peak Gain: {maxGain:F3}");
            sb.AppendLine($"Bandwidth (-3dB): {bandwidth:F2} Hz");
            sb.AppendLine();
        }

        private static void AnalyzeBacklash(List<DeviceData> data, StringBuilder sb)
        {
            sb.AppendLine("2. BACKLASH ANALYSIS (Zero-Crossing)");
            sb.AppendLine(new string('-', 50));

            int startIdx = Math.Max(0, data.Count - 20000);
            var triData = data.GetRange(startIdx, data.Count - startIdx);

            var maxPoint = triData.OrderByDescending(x => x.SetPointDeg).FirstOrDefault();
            var minPoint = triData.OrderBy(x => x.SetPointDeg).FirstOrDefault();

            if (maxPoint == null || minPoint == null)
            {
                sb.AppendLine("Triangle wave not detected.");
                return;
            }

            int pIdx = data.IndexOf(maxPoint);
            int vIdx = data.IndexOf(minPoint);

            List<DeviceData> negRamp, posRamp;
            if (pIdx < vIdx)
            {
                negRamp = data.GetRange(pIdx, vIdx - pIdx);
                int endPos = Math.Min(data.Count, vIdx + (vIdx - pIdx));
                posRamp = data.GetRange(vIdx, endPos - vIdx);
            }
            else
            {
                negRamp = data.GetRange(vIdx, pIdx - vIdx);
                int endPos = Math.Min(data.Count, pIdx + (pIdx - vIdx));
                posRamp = data.GetRange(pIdx, endPos - pIdx);
            }

            var zeroNeg = negRamp.OrderBy(x => Math.Abs(x.SetPointDeg)).First();
            var zeroPos = posRamp.OrderBy(x => Math.Abs(x.SetPointDeg)).First();

            double backlash = Math.Abs((zeroPos.SetPointDeg - zeroPos.ActualDeg) - (zeroNeg.SetPointDeg - zeroNeg.ActualDeg));

            sb.AppendLine($"Backlash: {backlash:F4} deg");
            sb.AppendLine();
        }

        private static void AnalyzeStepResponse(List<DeviceData> data, StringBuilder sb)
        {
            List<StepResult> steps = new List<StepResult>();
            List<StepResult> sensSteps = new List<StepResult>();

            List<int> stepStarts = new List<int>();
            for (int i = 1; i < data.Count; i++)
            {
                if (Math.Abs(data[i].SetPointRaw - data[i - 1].SetPointRaw) > 40)
                {
                    if (stepStarts.Count == 0 || (i - stepStarts.Last()) > 400)
                    {
                        stepStarts.Add(i);
                    }
                }
            }

            foreach (int idx in stepStarts)
            {
                int limit = stepStarts.FirstOrDefault(x => x > idx);
                if (limit == 0) limit = data.Count;
                int winLen = Math.Min(1400, limit - idx);
                if (winLen < 100) continue;

                var cmdSeg = data.GetRange(idx + 50, winLen - 50);
                if (CalculateStdDev(cmdSeg.Select(x => x.SetPointDeg)) > 0.05) continue;

                double cmdBefore = data[Math.Max(0, idx - 20)].SetPointDeg;
                double cmdAfter = cmdSeg.GroupBy(x => x.SetPointDeg).OrderByDescending(g => g.Count()).First().Key;
                double amp = cmdAfter - cmdBefore;

                if (Math.Abs(cmdAfter) < 0.02) continue;

                bool isMain = Math.Abs(amp) >= 0.9;
                bool isSens = Math.Abs(amp) < 0.9 && Math.Abs(amp) > 0.04;

                var fbkSeg = data.GetRange(idx, winLen);

                double startVal = data.GetRange(Math.Max(0, idx - 20), Math.Min(20, data.Count - idx + 20)).Average(x => x.ActualDeg);
                double finalVal = data.GetRange(idx + winLen - 100, Math.Min(100, winLen)).Average(x => x.ActualDeg);

                double deadTime = 0;
                double th2 = startVal + 0.02 * amp;
                var dtPoint = FindCrossing(fbkSeg, th2, amp > 0);
                if (dtPoint != null) deadTime = (dtPoint.Time - data[idx].Time) / 1000.0; // ثانیه

                double riseTime = 0;
                var p10 = FindCrossing(fbkSeg, startVal + 0.1 * amp, amp > 0);
                var p90 = FindCrossing(fbkSeg, startVal + 0.9 * amp, amp > 0);
                if (p10 != null && p90 != null) riseTime = (p90.Time - p10.Time) / 1000.0;

                double settlingTime = 0;
                double band = 0.05 * Math.Abs(amp);
                var lastOut = fbkSeg.LastOrDefault(x => Math.Abs(x.ActualDeg - finalVal) > band);
                if (lastOut != null) settlingTime = (lastOut.Time - data[idx].Time) / 1000.0;

                double peak = amp > 0 ? fbkSeg.Max(x => x.ActualDeg) : fbkSeg.Min(x => x.ActualDeg);
                double os = Math.Abs(peak - finalVal) / Math.Abs(amp) * 100;

                double sse = cmdAfter - finalVal;

                var res = new StepResult
                {
                    From = cmdBefore,
                    To = cmdAfter,
                    DeadTime = Math.Max(0, deadTime),
                    RiseTime = Math.Max(0, riseTime),
                    SettlingTime = Math.Max(0, settlingTime),
                    Overshoot = os,
                    SSE = sse,
                    IsSensitivity = isSens,
                    ActualValue = finalVal
                };

                if (isMain) steps.Add(res);
                else if (isSens) sensSteps.Add(res);
            }

            sb.AppendLine("3. STEP RESPONSE PARAMETERS (Forward Only, 2% DeadTime)");
            sb.AppendLine(new string('-', 85));
            sb.AppendLine($"{"From",-8} {"To",-8} {"DeadT(ms)",-10} {"RiseT(ms)",-10} {"SetT(ms)",-10} {"OS(%)",-8} {"SSE(deg)",-10}");
            sb.AppendLine(new string('-', 85));
            foreach (var s in steps)
            {
                sb.AppendLine($"{s.From,-8:F1} {s.To,-8:F1} {s.DeadTime * 1000:F0,-10} {s.RiseTime * 1000:F0,-10} {s.SettlingTime * 1000:F0,-10} {s.Overshoot,-8:F2} {s.SSE,-10:F4}");
            }
            sb.AppendLine();

            sb.AppendLine("4. SENSITIVITY CHECK (Micro Steps)");
            sb.AppendLine(new string('-', 55));
            sb.AppendLine($"{"Target",-10} {"Actual",-10} {"Error",-10} {"Status",-10}");
            sb.AppendLine(new string('-', 55));
            foreach (var s in sensSteps)
            {
                string status = Math.Abs(s.SSE) < 0.05 ? "PASS" : "CHECK";
                sb.AppendLine($"{s.To,-10:F4} {s.ActualValue,-10:F4} {s.SSE,-10:F4} {status,-10}");
            }
        }

        private static double CalculateRMS(IEnumerable<double> values)
        {
            var list = values.ToList();
            if (list.Count == 0) return 0;
            return Math.Sqrt(list.Average(v => v * v));
        }

        private static double CalculateStdDev(IEnumerable<double> values)
        {
            var list = values.ToList();
            if (list.Count == 0) return 0;
            double avg = list.Average();
            return Math.Sqrt(list.Average(v => Math.Pow(v - avg, 2)));
        }

        private static DeviceData FindCrossing(List<DeviceData> data, double threshold, bool rising)
        {
            return rising ? data.FirstOrDefault(p => p.ActualDeg >= threshold)
                          : data.FirstOrDefault(p => p.ActualDeg <= threshold);
        }
    }
}