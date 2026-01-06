using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;

namespace DeviceAnalisys_v5
{
    internal class ExcelOperation
    {
        /// <summary>
        /// Loads data from Excel file and adds it to DBList for plotting.
        /// Supports files with both Raw and Deg columns, or older formats.
        /// </summary>
        public void LoadExcelToDiagramQueue(string filePath)
        {
            XLWorkbook workbook = null;
            try
            {
                workbook = new XLWorkbook(filePath);

                foreach (var worksheet in workbook.Worksheets)
                {
                    // Skip non-device sheets
                    if (worksheet.Name == "Summary" ||
                        (!worksheet.Name.StartsWith("Device_") && !worksheet.Name.StartsWith("Dev_")))
                        continue;

                    int row = 2; // Row 1 is header
                    while (!worksheet.Cell(row, 1).IsEmpty())
                    {
                        if (!worksheet.Cell(row, 1).TryGetValue(out int testId))
                            break;

                        string serialNumber = string.Empty;
                        int time = 0;

                        // Try to read Raw and Deg (new format with 11 columns)
                        double setPointRaw = 0, actualRaw = 0, pitchRaw = 0, rollRaw = 0;
                        double setPointDeg = 0, actualDeg = 0, pitchDeg = 0, rollDeg = 0;

                        worksheet.Cell(row, 2).TryGetValue(out serialNumber);
                        worksheet.Cell(row, 3).TryGetValue(out time);

                        // New format: Raw + Deg
                        if (worksheet.Row(1).CellsUsed().Count() >= 11)
                        {
                            worksheet.Cell(row, 4).TryGetValue(out setPointRaw);
                            worksheet.Cell(row, 5).TryGetValue(out actualRaw);
                            worksheet.Cell(row, 6).TryGetValue(out pitchRaw);
                            worksheet.Cell(row, 7).TryGetValue(out rollRaw);
                            worksheet.Cell(row, 8).TryGetValue(out setPointDeg);
                            worksheet.Cell(row, 9).TryGetValue(out actualDeg);
                            worksheet.Cell(row, 10).TryGetValue(out pitchDeg);
                            worksheet.Cell(row, 11).TryGetValue(out rollDeg);
                        }
                        // Old format: only Deg
                        else
                        {
                            worksheet.Cell(row, 4).TryGetValue(out setPointDeg);
                            worksheet.Cell(row, 5).TryGetValue(out actualDeg);
                            worksheet.Cell(row, 6).TryGetValue(out pitchDeg);
                            worksheet.Cell(row, 7).TryGetValue(out rollDeg);

                            // Calculate Raw from Deg if needed (for compatibility)
                            setPointRaw = setPointDeg * GlobalData.Scale;
                            actualRaw = actualDeg * GlobalData.Scale;
                            pitchRaw = pitchDeg * GlobalData.Scale;
                            rollRaw = rollDeg * GlobalData.Scale;
                        }

                        var data = new DeviceData
                        {
                            TestID = testId,
                            Time = time,
                            SetPointRaw = setPointRaw,
                            ActualRaw = actualRaw,
                            PitchRaw = pitchRaw,
                            RollRaw = rollRaw,
                            SetPointDeg = setPointDeg,
                            ActualDeg = actualDeg,
                            PitchDeg = pitchDeg,
                            RollDeg = rollDeg,
                            SerialNumber = string.IsNullOrWhiteSpace(serialNumber)
                                ? worksheet.Name.Replace("Device_", "SN-DEV")
                                : serialNumber.Trim()
                        };

                        // Final fallback for serial number
                        if (string.IsNullOrWhiteSpace(data.SerialNumber) ||
                            data.SerialNumber.StartsWith("SN-DEV") ||
                            data.SerialNumber.StartsWith("SN-UNKNOWN"))
                        {
                            int devIndex = testId - 1;
                            if (devIndex >= 0 && devIndex < MainWindow.CurrentDeviceSerials.Length)
                            {
                                data.SerialNumber = MainWindow.CurrentDeviceSerials[devIndex];
                            }
                        }

                        GlobalData.DBList.Add(data);
                        row++;
                    }
                }

                MessageBox.Show("Excel file loaded successfully.\nData is ready for plotting.",
                    "Load Complete", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading Excel file:\n{ex.Message}",
                    "Load Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                workbook?.Dispose();
            }
        }

        /// <summary>
        /// Reads all device data from Excel file, grouped by sheet/device.
        /// </summary>
        public Dictionary<string, List<DeviceData>> ReadAllDevicesFromExcel(string filePath)
        {
            var result = new Dictionary<string, List<DeviceData>>();
            XLWorkbook workbook = null;
            try
            {
                workbook = new XLWorkbook(filePath);

                foreach (var worksheet in workbook.Worksheets)
                {
                    if (worksheet.Name == "Summary" ||
                        (!worksheet.Name.StartsWith("Device_") && !worksheet.Name.StartsWith("Dev_")))
                        continue;

                    string deviceKey = worksheet.Name;
                    var deviceData = new List<DeviceData>();
                    int row = 2;

                    while (!worksheet.Cell(row, 1).IsEmpty())
                    {
                        if (!worksheet.Cell(row, 1).TryGetValue(out int testId))
                            break;

                        string serial = string.Empty;
                        int time = 0;
                        double setPointRaw = 0, actualRaw = 0, pitchRaw = 0, rollRaw = 0;
                        double setPointDeg = 0, actualDeg = 0, pitchDeg = 0, rollDeg = 0;

                        worksheet.Cell(row, 2).TryGetValue(out serial);
                        worksheet.Cell(row, 3).TryGetValue(out time);

                        if (worksheet.Row(1).CellsUsed().Count() >= 11)
                        {
                            worksheet.Cell(row, 4).TryGetValue(out setPointRaw);
                            worksheet.Cell(row, 5).TryGetValue(out actualRaw);
                            worksheet.Cell(row, 6).TryGetValue(out pitchRaw);
                            worksheet.Cell(row, 7).TryGetValue(out rollRaw);
                            worksheet.Cell(row, 8).TryGetValue(out setPointDeg);
                            worksheet.Cell(row, 9).TryGetValue(out actualDeg);
                            worksheet.Cell(row, 10).TryGetValue(out pitchDeg);
                            worksheet.Cell(row, 11).TryGetValue(out rollDeg);
                        }
                        else
                        {
                            worksheet.Cell(row, 4).TryGetValue(out setPointDeg);
                            worksheet.Cell(row, 5).TryGetValue(out actualDeg);
                            worksheet.Cell(row, 6).TryGetValue(out pitchDeg);
                            worksheet.Cell(row, 7).TryGetValue(out rollDeg);

                            setPointRaw = setPointDeg * GlobalData.Scale;
                            actualRaw = actualDeg * GlobalData.Scale;
                            pitchRaw = pitchDeg * GlobalData.Scale;
                            rollRaw = rollDeg * GlobalData.Scale;
                        }

                        var dataPoint = new DeviceData
                        {
                            TestID = testId,
                            Time = time,
                            SetPointRaw = setPointRaw,
                            ActualRaw = actualRaw,
                            PitchRaw = pitchRaw,
                            RollRaw = rollRaw,
                            SetPointDeg = setPointDeg,
                            ActualDeg = actualDeg,
                            PitchDeg = pitchDeg,
                            RollDeg = rollDeg,
                            SerialNumber = string.IsNullOrWhiteSpace(serial) ? deviceKey : serial.Trim()
                        };

                        deviceData.Add(dataPoint);
                        row++;
                    }

                    if (deviceData.Count > 0)
                    {
                        deviceData.Sort((a, b) => a.Time.CompareTo(b.Time));
                        result[deviceKey] = deviceData;
                    }
                }

                return result;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to read Excel file:\n{ex.Message}",
                    "Read Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return new Dictionary<string, List<DeviceData>>();
            }
            finally
            {
                workbook?.Dispose();
            }
        }

        /// <summary>
        /// Saves all collected data to Excel file, with each device in its own sheet.
        /// Stores both Raw and Deg values.
        /// </summary>
        public void SaveToExcel(string filePath = null)
        {
            try
            {
                if (GlobalData.DBList == null || GlobalData.DBList.Count == 0)
                {
                    MessageBox.Show("No data available to save.", "No Data",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                if (string.IsNullOrEmpty(filePath))
                {
                    var saveDialog = new Microsoft.Win32.SaveFileDialog
                    {
                        Filter = "Excel Files (*.xlsx)|*.xlsx",
                        FileName = $"TestData_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx",
                        DefaultExt = ".xlsx"
                    };
                    if (saveDialog.ShowDialog() != true)
                        return;
                    filePath = saveDialog.FileName;
                }

                using (var workbook = new XLWorkbook())
                {
                    var groupedData = GlobalData.DBList
                        .GroupBy(d => d.TestID)
                        .OrderBy(g => g.Key);

                    foreach (var group in groupedData)
                    {
                        int deviceIndex = group.Key - 1;
                        string sheetName = $"Device_{group.Key}";

                        var firstValidSerial = group.FirstOrDefault(d =>
                            !string.IsNullOrWhiteSpace(d.SerialNumber) &&
                            !d.SerialNumber.StartsWith("SN-DEV") &&
                            !d.SerialNumber.StartsWith("SN-UNKNOWN"));

                        if (firstValidSerial != null)
                        {
                            string cleanSerial = firstValidSerial.SerialNumber.Trim()
                                .Replace("/", "-").Replace("\\", "-")
                                .Replace("?", "").Replace("*", "")
                                .Replace("[", "(").Replace("]", ")");

                            sheetName = $"Dev_{group.Key}_{cleanSerial}";
                            if (sheetName.Length > 31)
                                sheetName = sheetName.Substring(0, 31);
                        }

                        var ws = workbook.Worksheets.Add(sheetName);

                        // Headers (Raw + Deg)
                        ws.Cell(1, 1).Value = "TestID";
                        ws.Cell(1, 2).Value = "Serial Number";
                        ws.Cell(1, 3).Value = "Time";
                        ws.Cell(1, 4).Value = "SetPointRaw";
                        ws.Cell(1, 5).Value = "ActualRaw";
                        ws.Cell(1, 6).Value = "PitchRaw";
                        ws.Cell(1, 7).Value = "RollRaw";
                        ws.Cell(1, 8).Value = "SetPointDeg";
                        ws.Cell(1, 9).Value = "ActualDeg";
                        ws.Cell(1, 10).Value = "PitchDeg";
                        ws.Cell(1, 11).Value = "RollDeg";

                        var headerRange = ws.Range("A1:K1");
                        headerRange.Style.Font.Bold = true;
                        headerRange.Style.Fill.BackgroundColor = XLColor.FromArgb(0, 122, 204);
                        headerRange.Style.Font.FontColor = XLColor.White;
                        headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                        // Data rows
                        int row = 2;
                        foreach (var data in group.OrderBy(d => d.Time))
                        {
                            ws.Cell(row, 1).Value = data.TestID;
                            ws.Cell(row, 2).Value = data.SerialNumber;
                            ws.Cell(row, 3).Value = data.Time;
                            ws.Cell(row, 4).Value = data.SetPointRaw;
                            ws.Cell(row, 5).Value = data.ActualRaw;
                            ws.Cell(row, 6).Value = data.PitchRaw;
                            ws.Cell(row, 7).Value = data.RollRaw;
                            ws.Cell(row, 8).Value = data.SetPointDeg;
                            ws.Cell(row, 9).Value = data.ActualDeg;
                            ws.Cell(row, 10).Value = data.PitchDeg;
                            ws.Cell(row, 11).Value = data.RollDeg;
                            row++;
                        }

                        ws.Columns("A:K").AdjustToContents();
                    }

                    // Summary sheet
                    var summarySheet = workbook.Worksheets.Add("Summary", 1);
                    summarySheet.Cell(1, 1).Value = "Data Summary";
                    summarySheet.Cell(2, 1).Value = "Total points";
                    summarySheet.Cell(2, 2).Value = GlobalData.DBList.Count;
                    summarySheet.Cell(3, 1).Value = "Number of devices";
                    summarySheet.Cell(3, 2).Value = groupedData.Count();
                    summarySheet.Cell(5, 1).Value = "Generated at";
                    summarySheet.Cell(5, 2).Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    summarySheet.Columns("A:B").AdjustToContents();

                    workbook.SaveAs(filePath);
                }

                MessageBox.Show($"Data successfully saved!\n{filePath}\nEach device is saved in its own sheet.",
                    "Save Completed", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error while saving Excel file:\n{ex.Message}",
                    "Save Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}