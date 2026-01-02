using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace DeviceAnalisys_v5
{
    internal class ExcelOperation
    {
        public void LoadExcelToDiagramQueue(string filePath)
        {
            try
            {
                using (var workbook = new XLWorkbook(filePath))
                {
                    var ws = workbook.Worksheet(1);

                    int headerColumnCount = ws.Row(1).CellsUsed().Count();

                    bool hasSerialColumn = headerColumnCount >= 7;

                    int row = 2;
                    while (!ws.Cell(row, 1).IsEmpty())
                    {
                        if (!ws.Cell(row, 1).TryGetValue(out int testId))
                            break;

                        string serialNumber = "";
                        int time = 0;
                        double setPoint = 0, actual = 0, pitch = 0, roll = 0;

                        if (hasSerialColumn)
                        {
                            ws.Cell(row, 2).TryGetValue(out serialNumber);
                            ws.Cell(row, 3).TryGetValue(out time);
                            ws.Cell(row, 4).TryGetValue(out setPoint);
                            ws.Cell(row, 5).TryGetValue(out actual);
                            ws.Cell(row, 6).TryGetValue(out pitch);
                            ws.Cell(row, 7).TryGetValue(out roll);
                        }
                        else
                        {
                            ws.Cell(row, 2).TryGetValue(out time);
                            ws.Cell(row, 3).TryGetValue(out setPoint);
                            ws.Cell(row, 4).TryGetValue(out actual);
                            ws.Cell(row, 5).TryGetValue(out pitch);
                            ws.Cell(row, 6).TryGetValue(out roll);

                            int deviceIndex = testId - 1;
                            if (deviceIndex >= 0 && deviceIndex < MainWindow.CurrentDeviceSerials.Length)
                            {
                                serialNumber = MainWindow.CurrentDeviceSerials[deviceIndex];
                            }
                            else
                            {
                                serialNumber = $"SN-UNKNOWN-{testId}";
                            }
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

                        // همیشه سریال رو قبل از enqueue چک و ست کن
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
                }

                MessageBox.Show("Excel file loaded successfully. Data is ready for plotting.",
                  "Load Complete", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading Excel file:\n{ex.Message}",
                          "Load Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        public void SaveToExcel(string filePath = null)
        {
            try
            {
                if (string.IsNullOrEmpty(filePath))
                {
                    var saveDialog = new Microsoft.Win32.SaveFileDialog
                    {
                        Filter = "Excel Files (*.xlsx)|*.xlsx",
                        FileName = "TestData_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx",
                        DefaultExt = ".xlsx"
                    };

                    if (saveDialog.ShowDialog() != true)
                        return; 

                    filePath = saveDialog.FileName;
                }

                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("Test Data");

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

                    int row = 2;
                    foreach (var data in GlobalData.DBList)
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

                    workbook.SaveAs(filePath);

                    MessageBox.Show($"Data successfully saved to Excel!\n{filePath}",
                                    "Save Completed", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error saving to Excel:\n{ex.Message}",
                                "Save Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
