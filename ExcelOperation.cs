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
        /// Loads data from Excel file and enqueues it for diagram/plotting.
        /// Supports both formats: with serial number column and without it.
        /// </summary>
        public void LoadExcelToDiagramQueue(string filePath)
        {
            XLWorkbook workbook = null;
            try
            {
                workbook = new XLWorkbook(filePath);
                var worksheet = workbook.Worksheet(1);

                int headerColumnCount = worksheet.Row(1).CellsUsed().Count();
                bool hasSerialColumn = headerColumnCount >= 7;

                int row = 2;
                while (!worksheet.Cell(row, 1).IsEmpty())
                {
                    if (!worksheet.Cell(row, 1).TryGetValue(out int testId))
                        break;

                    string serialNumber = string.Empty;
                    int time = 0;
                    double setPoint = 0, actual = 0, pitch = 0, roll = 0;

                    if (hasSerialColumn)
                    {
                        worksheet.Cell(row, 2).TryGetValue(out serialNumber);
                        worksheet.Cell(row, 3).TryGetValue(out time);
                        worksheet.Cell(row, 4).TryGetValue(out setPoint);
                        worksheet.Cell(row, 5).TryGetValue(out actual);
                        worksheet.Cell(row, 6).TryGetValue(out pitch);
                        worksheet.Cell(row, 7).TryGetValue(out roll);
                    }
                    else
                    {
                        worksheet.Cell(row, 2).TryGetValue(out time);
                        worksheet.Cell(row, 3).TryGetValue(out setPoint);
                        worksheet.Cell(row, 4).TryGetValue(out actual);
                        worksheet.Cell(row, 5).TryGetValue(out pitch);
                        worksheet.Cell(row, 6).TryGetValue(out roll);

                        int deviceIndex = testId - 1;
                        serialNumber = (deviceIndex >= 0 && deviceIndex < MainWindow.CurrentDeviceSerials.Length)
                            ? MainWindow.CurrentDeviceSerials[deviceIndex]
                            : $"SN-UNKNOWN-{testId}";
                    }

                    var data = new DeviceData
                    {
                        TestID = testId,
                        Time = time,
                        SetPoint = setPoint,
                        Actual = actual,
                        Pitch = pitch,
                        Roll = roll,
                        SerialNumber = string.IsNullOrWhiteSpace(serialNumber)
                            ? $"SN-DEV{testId:D3}"
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

                    GlobalData.DiagramQueue.Enqueue(data);
                    row++;
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
        /// Suitable for plotting multiple series (one series per device/sheet).
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
                    // Skip non-device sheets (Summary or other irrelevant sheets)
                    if (worksheet.Name == "Summary" ||
                        (!worksheet.Name.StartsWith("Device_") && !worksheet.Name.StartsWith("Dev_")))
                        continue;

                    string deviceKey = worksheet.Name;
                    var deviceData = new List<DeviceData>();

                    int row = 2; // Row 1 is assumed to be header
                    while (!worksheet.Cell(row, 1).IsEmpty())
                    {
                        if (!worksheet.Cell(row, 1).TryGetValue(out int testId))
                            break;

                        worksheet.Cell(row, 2).TryGetValue(out string serial);
                        worksheet.Cell(row, 3).TryGetValue(out int time);
                        worksheet.Cell(row, 4).TryGetValue(out double setPoint);
                        worksheet.Cell(row, 5).TryGetValue(out double actual);
                        worksheet.Cell(row, 6).TryGetValue(out double pitch);
                        worksheet.Cell(row, 7).TryGetValue(out double roll);

                        var dataPoint = new DeviceData
                        {
                            TestID = testId,
                            SerialNumber = string.IsNullOrWhiteSpace(serial) ? deviceKey : serial.Trim(),
                            Time = time,
                            SetPoint = setPoint,
                            Actual = actual,
                            Pitch = pitch,
                            Roll = roll
                        };

                        deviceData.Add(dataPoint);
                        row++;
                    }

                    if (deviceData.Count > 0)
                    {
                        // Sort data by time (just in case)
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

                        // Headers
                        ws.Cell(1, 1).Value = "TestID";
                        ws.Cell(1, 2).Value = "Serial Number";
                        ws.Cell(1, 3).Value = "Time";
                        ws.Cell(1, 4).Value = "SetPoint";
                        ws.Cell(1, 5).Value = "Actual";
                        ws.Cell(1, 6).Value = "Pitch";
                        ws.Cell(1, 7).Value = "Roll";

                        var headerRange = ws.Range("A1:G1");
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
                            ws.Cell(row, 4).Value = data.SetPoint;
                            ws.Cell(row, 5).Value = data.Actual;
                            ws.Cell(row, 6).Value = data.Pitch;
                            ws.Cell(row, 7).Value = data.Roll;
                            row++;
                        }

                        ws.Columns("A:G").AdjustToContents();
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