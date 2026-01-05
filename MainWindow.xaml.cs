using System;
using System.Collections.Generic;
using System.IO.Ports;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using Microsoft.Win32;
using ScottPlot;
using ScottPlot.Plottables;
using ScottPlot.WPF;
using ClosedXML.Excel;
using System.Threading;
using System.Windows.Media.Animation;

namespace DeviceAnalisys_v5
{
    public partial class MainWindow : Window
    {
        private readonly ExcelOperation excelOperation = new ExcelOperation();
        private SerialPortReader serialPortReader;
        private DispatcherTimer plotTimer;

        private readonly int deviceCount = 4;
        private readonly string[] parameters = { "SetPoint", "Actual", "Pitch", "Roll" };

        private List<DataLogger>[] loggerLists;
        private Marker[] highlightMarkers;
        private WpfPlot[] plots;
        private TextBlock[] liveValueTexts;

        private string[] deviceSerialNumbers = new string[4];
        public static string[] CurrentDeviceSerials { get; private set; } = new string[4];

        private bool isPaused = false;
        private int totalPoints = 0;
        private double lastTimeIndex = -1;

        // Track fullscreen and HD modes separately
        private int? fullScreenDevice = null;
        private int? hdModeDevice = null;

        private ColumnDefinition leftPanelColumn;

        public MainWindow()
        {
            InitializeComponent();

            // Reference to left panel column for hiding/showing
            leftPanelColumn = mainContentGrid.ColumnDefinitions[0];

            PlotInitialSetup();
            LoadAvailablePorts();

            deviceSerialNumbers[0] = "SN-DEV001";
            deviceSerialNumbers[1] = "SN-DEV002";
            deviceSerialNumbers[2] = "SN-DEV003";
            deviceSerialNumbers[3] = "SN-DEV004";

            CurrentDeviceSerials = (string[])deviceSerialNumbers.Clone();
        }

        private void LoadAvailablePorts()
        {
            string[] ports = SerialPort.GetPortNames();
            cmbComPorts.Items.Clear();

            if (ports.Length == 0)
            {
                cmbComPorts.Items.Add("No ports found");
                cmbComPorts.SelectedIndex = 0;
                btnConnect.IsEnabled = false;
            }
            else
            {
                foreach (string port in ports.OrderBy(p => p))
                {
                    cmbComPorts.Items.Add(port);
                }

                if (ports.Contains("COM3"))
                {
                    cmbComPorts.SelectedItem = "COM3";
                }
                else if (ports.Length > 0)
                {
                    cmbComPorts.SelectedIndex = 0;
                }

                btnConnect.IsEnabled = true;
            }
        }

        private void RefreshPorts_Click(object sender, RoutedEventArgs e)
        {
            LoadAvailablePorts();
        }

        private void Connect_Click(object sender, RoutedEventArgs e)
        {
            if (cmbComPorts.SelectedItem == null || cmbComPorts.SelectedItem.ToString() == "No ports found")
            {
                MessageBox.Show("No COM port selected or available.", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            string selectedPort = cmbComPorts.SelectedItem.ToString();

            try
            {
                serialPortReader = new SerialPortReader(selectedPort);
                serialPortReader.Start();

                txtConnectionStatus.Text = $"Connected to {selectedPort}";
                txtConnectionStatus.Foreground = new SolidColorBrush(System.Windows.Media.Colors.LimeGreen);

                btnConnect.IsEnabled = false;
                btnDisconnect.IsEnabled = true;
                cmbComPorts.IsEnabled = false;

                statusText.Text = $"Connected to {selectedPort} | Waiting for data...";

                if (isPaused)
                {
                    PauseResume_Click(null, null);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Cannot open port {selectedPort}:\n{ex.Message}", "Connection Error", MessageBoxButton.OK, MessageBoxImage.Error);
                txtConnectionStatus.Text = "Connection failed";
                txtConnectionStatus.Foreground = new SolidColorBrush(System.Windows.Media.Colors.Red);
            }
        }

        private void Disconnect_Click(object sender, RoutedEventArgs e)
        {
            serialPortReader?.Stop();

            txtConnectionStatus.Text = "Disconnected";
            txtConnectionStatus.Foreground = new SolidColorBrush(System.Windows.Media.Colors.OrangeRed);

            btnConnect.IsEnabled = true;
            btnDisconnect.IsEnabled = false;
            cmbComPorts.IsEnabled = true;

            statusText.Text = "Disconnected from serial port";
            txtTestStatus.Text = "Test Status: Idle";
            txtTestStatus.Foreground = new SolidColorBrush(System.Windows.Media.Colors.LightGray);
        }

        private void PlotInitialSetup()
        {
            plots = new WpfPlot[] { plot0, plot1, plot2, plot3 };
            loggerLists = new List<DataLogger>[deviceCount];
            highlightMarkers = new Marker[deviceCount];

            liveValueTexts = new TextBlock[]
            {
                valSetPoint0, valActual0, valPitch0, valRoll0,
                valSetPoint1, valActual1, valPitch1, valRoll1,
                valSetPoint2, valActual2, valPitch2, valRoll2,
                valSetPoint3, valActual3, valPitch3, valRoll3
            };

            for (int i = 0; i < deviceCount; i++)
            {
                loggerLists[i] = new List<DataLogger>();

                var plt = plots[i].Plot;
                plt.FigureBackground.Color = ScottPlot.Color.FromHex("#1E1E1E");
                plt.DataBackground.Color = ScottPlot.Color.FromHex("#1E1E1E");

                plt.Axes.Title.Label.Text = $"Device {i + 1} - {deviceSerialNumbers[i]}";
                plt.Axes.Title.Label.ForeColor = ScottPlot.Colors.White;
                plt.Axes.Title.Label.FontSize = 16;

                plt.Axes.Bottom.Label.Text = "Index";
                plt.Axes.Left.Label.Text = "Value";

                plt.Legend.IsVisible = true;
                plt.Legend.Alignment = Alignment.UpperRight;
                plt.Legend.FontSize = 12;

                plots[i].Menu = null;

                plt.Axes.Bottom.Label.ForeColor = ScottPlot.Colors.White;
                plt.Axes.Left.Label.ForeColor = ScottPlot.Colors.White;
                plt.Axes.Bottom.TickLabelStyle.ForeColor = ScottPlot.Colors.White;
                plt.Axes.Left.TickLabelStyle.ForeColor = ScottPlot.Colors.White;

                plt.Axes.Bottom.TickLabelStyle.FontSize = 11;
                plt.Axes.Left.TickLabelStyle.FontSize = 11;

                plt.Grid.IsVisible = false;
                plt.Axes.Color(ScottPlot.Colors.LightGray);

                highlightMarkers[i] = plt.Add.Marker(0, 0);
                highlightMarkers[i].IsVisible = false;
                highlightMarkers[i].Size = 12f;
                highlightMarkers[i].Shape = MarkerShape.FilledCircle;
                highlightMarkers[i].Color = ScottPlot.Color.FromHex("#FFFF00");

                plots[i].MouseMove += WpfPlot_MouseMove;
                plots[i].MouseLeave += WpfPlot_MouseLeave;

                for (int j = 0; j < parameters.Length; j++)
                {
                    var logger = plt.Add.DataLogger();
                    logger.LegendText = parameters[j];
                    logger.LineWidth = 3.0f;
                    logger.MarkerSize = 0;
                    logger.ManageAxisLimits = false;

                    var palette = new ScottPlot.Palettes.Category10();
                    logger.Color = palette.GetColor(j);

                    loggerLists[i].Add(logger);
                }

                loggerLists[i][2].IsVisible = false; // Pitch hidden by default
                loggerLists[i][3].IsVisible = false; // Roll hidden by default
            }
            plotTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(100) };
            plotTimer.Tick += PlotTimer_Tick;
            plotTimer.Start();
            cmbFps.SelectionChanged += cmbFps_SelectionChanged;
        }

        private void PlotTimer_Tick(object sender, EventArgs e)
        {
            // Skip processing if plotting is paused
            if (isPaused) return;

            // Constants for batch processing and maximum points per plot
            const int MAX_PER_TICK = 1500;
            const int MAX_POINTS_PER_PLOT = 35000;

            bool updated = false;
            var updatedPlots = new HashSet<int>();
            double currentMaxTime = lastTimeIndex;
            int processed = 0;

            // Process data from the queue in batches
            while (processed < MAX_PER_TICK && GlobalData.DiagramQueue.TryDequeue(out DeviceData data))
            {
                int deviceIndex = data.TestID - 1;
                if (deviceIndex < 0 || deviceIndex >= deviceCount) continue;

                // Fallback for missing or invalid serial number
                if (string.IsNullOrWhiteSpace(data.SerialNumber) ||
                    data.SerialNumber.StartsWith("SN-DEV") ||
                    data.SerialNumber.StartsWith("SN-UNKNOWN") ||
                    data.SerialNumber == $"SN-DEV{data.TestID:D3}")
                {
                    data.SerialNumber = deviceSerialNumbers[deviceIndex];
                }

                // Store data in global list
                GlobalData.DBList.Add(data);

                // Add data to corresponding plot loggers
                var loggerList = loggerLists[deviceIndex];
                loggerList[0].Add(data.Time, data.SetPoint);
                loggerList[1].Add(data.Time, data.Actual);
                loggerList[2].Add(data.Time, data.Pitch);
                loggerList[3].Add(data.Time, data.Roll);

                // Trim old points if exceeding maximum limit (to prevent memory issues)
                foreach (var logger in loggerList)
                {
                    if (logger.Data.Coordinates.Count > MAX_POINTS_PER_PLOT)
                    {
                        int excess = logger.Data.Coordinates.Count - MAX_POINTS_PER_PLOT;
                        logger.Data.Coordinates.RemoveRange(0, excess);
                    }
                }

                // Update live value displays for this device
                int baseIndex = deviceIndex * 4;
                liveValueTexts[baseIndex + 0].Text = $"SetPoint: {data.SetPoint:F3}";
                liveValueTexts[baseIndex + 1].Text = $"Actual: {data.Actual:F3}";
                liveValueTexts[baseIndex + 2].Text = $"Pitch: {data.Pitch:F3}";
                liveValueTexts[baseIndex + 3].Text = $"Roll: {data.Roll:F3}";

                totalPoints++;
                updated = true;
                updatedPlots.Add(deviceIndex);

                // Update the latest time index
                if (data.Time > currentMaxTime)
                    currentMaxTime = data.Time;

                processed++;
            }

            // Refresh UI if any new data was processed
            if (updated)
            {
                foreach (int index in updatedPlots)
                {
                    plots[index].Plot.Axes.AutoScale();
                    plots[index].Refresh();
                }

                livePointsText.Text = totalPoints.ToString();
                liveLastUpdate.Text = currentMaxTime.ToString("F0");
                lastTimeIndex = currentMaxTime;

                statusText.Text = $"Running | Points: {totalPoints}";
            }

            // Check if loading from Excel/queue is finished
            // (no new data processed AND queue is empty AND we have some points)
            // فقط یک بار وقتی لود تمام شد، عنوان پلات‌ها رو آپدیت کن
            if (!updated && GlobalData.DiagramQueue.IsEmpty && totalPoints > 0)
            {
                statusText.Text = $"Finished loading | Total Points: {totalPoints}";
                statusText.Foreground = new SolidColorBrush(System.Windows.Media.Colors.LimeGreen);

                // آپدیت عنوان‌ها بر اساس سریال واقعی در داده‌ها
                bool titlesUpdated = false;
                for (int i = 0; i < deviceCount; i++)
                {
                    int testId = i + 1;
                    var anyData = GlobalData.DBList
                        .FirstOrDefault(d => d.TestID == testId && !string.IsNullOrWhiteSpace(d.SerialNumber));

                    if (anyData != null)
                    {
                        string realSerial = anyData.SerialNumber.Trim();
                        if (deviceSerialNumbers[i] != realSerial)
                        {
                            deviceSerialNumbers[i] = realSerial;
                            plots[i].Plot.Axes.Title.Label.Text = $"Device {i + 1} - {realSerial}";
                            plots[i].Refresh();
                            titlesUpdated = true;
                        }
                    }
                }

                if (titlesUpdated)
                {
                    CurrentDeviceSerials = (string[])deviceSerialNumbers.Clone();
                }
            }
        }

        private void WpfPlot_MouseMove(object sender, MouseEventArgs e)
        {
            if (!(sender is WpfPlot wpfPlot)) return;

            int plotIndex = Array.IndexOf(plots, wpfPlot);
            if (plotIndex < 0) return;

            var plt = wpfPlot.Plot;
            var loggerList = loggerLists[plotIndex];
            if (loggerList.Count == 0) return;

            Pixel mousePixel = new Pixel((float)e.GetPosition(wpfPlot).X, (float)e.GetPosition(wpfPlot).Y);
            Coordinates mouseCoords = plt.GetCoordinates(mousePixel);

            DataLogger closestLogger = null;
            Coordinates closestCoords = new Coordinates(double.NaN, double.NaN);
            double minDistanceSq = double.MaxValue;

            foreach (var logger in loggerList)
            {
                if (!logger.IsVisible) continue;

                var coords = logger.Data.Coordinates;
                for (int i = 0; i < coords.Count; i++)
                {
                    double dx = coords[i].X - mouseCoords.X;
                    double dy = coords[i].Y - mouseCoords.Y;
                    double distSq = dx * dx + dy * dy;

                    if (distSq < minDistanceSq)
                    {
                        minDistanceSq = distSq;
                        closestLogger = logger;
                        closestCoords = coords[i];
                    }
                }
            }

            Marker marker = highlightMarkers[plotIndex];

            if (closestLogger != null && Math.Sqrt(minDistanceSq) < 50)
            {
                string paramName = closestLogger.LegendText ?? "Unknown";
                double indexValue = closestCoords.X;
                double value = closestCoords.Y;

                plt.Axes.Title.Label.Text = $"{paramName} | Index: {indexValue:F1} | Value: {value:F3}";
                plt.Axes.Title.Label.ForeColor = ScottPlot.Colors.Cyan;
                plt.Axes.Title.Label.FontSize = 14;

                marker.Location = closestCoords;
                marker.IsVisible = true;
            }
            else
            {
                plt.Axes.Title.Label.Text = $"Device {plotIndex + 1} - {deviceSerialNumbers[plotIndex]}";
                plt.Axes.Title.Label.ForeColor = ScottPlot.Colors.White;
                plt.Axes.Title.Label.FontSize = 14;

                marker.IsVisible = false;
            }

            wpfPlot.Refresh();
        }

        private void WpfPlot_MouseLeave(object sender, MouseEventArgs e)
        {
            if (!(sender is WpfPlot wpfPlot)) return;

            int plotIndex = Array.IndexOf(plots, wpfPlot);
            if (plotIndex < 0) return;

            var plt = wpfPlot.Plot;

            plt.Axes.Title.Label.Text = $"Device {plotIndex + 1} - {deviceSerialNumbers[plotIndex]}";
            plt.Axes.Title.Label.ForeColor = ScottPlot.Colors.White;
            plt.Axes.Title.Label.FontSize = 14;

            highlightMarkers[plotIndex].IsVisible = false;

            wpfPlot.Refresh();
        }

        private void cmbFps_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (plotTimer == null) return;

            if (sender is ComboBox comboBox && comboBox.SelectedItem is ComboBoxItem selectedItem)
            {
                if (selectedItem.Tag != null && int.TryParse(selectedItem.Tag.ToString(), out int intervalMs))
                {
                    plotTimer.Interval = TimeSpan.FromMilliseconds(intervalMs);
                }
            }
        }

        private void PauseResume_Click(object sender, RoutedEventArgs e)
        {
            isPaused = !isPaused;

            if (isPaused)
            {
                plotTimer.Stop();
                ((Button)sender).Content = "Resume";
                statusText.Text = $"Paused | Points: {totalPoints}";
            }
            else
            {
                plotTimer.Start();
                ((Button)sender).Content = "Pause";
                statusText.Text = $"Running | Points: {totalPoints}";
            }
        }

        private async void ClearAll_Click(object sender, RoutedEventArgs e)
        {
            // Step 1: Disable button during animation & clear
            if (sender is Button clearButton)
            {
                clearButton.IsEnabled = false; // جلوگیری از کلیک چندباره

                // Animation: Scale up + slight fade
                var scaleTransform = new ScaleTransform(1, 1);
                var opacityAnim = new DoubleAnimation
                {
                    From = 1.0,
                    To = 0.7,
                    Duration = new Duration(TimeSpan.FromMilliseconds(150)),
                    AutoReverse = true
                };

                var scaleAnim = new DoubleAnimation
                {
                    From = 1.0,
                    To = 1.15,
                    Duration = new Duration(TimeSpan.FromMilliseconds(150)),
                    AutoReverse = true
                };

                clearButton.RenderTransform = scaleTransform;
                clearButton.RenderTransformOrigin = new Point(0.5, 0.5);

                scaleTransform.BeginAnimation(ScaleTransform.ScaleXProperty, scaleAnim);
                scaleTransform.BeginAnimation(ScaleTransform.ScaleYProperty, scaleAnim);
                clearButton.BeginAnimation(UIElement.OpacityProperty, opacityAnim);

                await Task.Delay(300); // صبر تا انیمیشن تموم بشه
            }

            // Step 2: Clear everything (your original code + full reset)
            foreach (var loggerList in loggerLists)
            {
                foreach (var logger in loggerList)
                {
                    logger.Data.Clear();
                }
            }

            totalPoints = 0;
            lastTimeIndex = -1;

            // Full memory reset (very important!)
            GlobalData.DBList.Clear();
            while (GlobalData.DiagramQueue.TryDequeue(out _)) { } // Empty the queue

            statusText.Text = "Cleared | Ready for new test | Points: 0";
            statusText.Foreground = new SolidColorBrush(System.Windows.Media.Colors.LimeGreen);

            livePointsText.Text = "0";
            liveLastUpdate.Text = "-";

            foreach (var tb in liveValueTexts)
            {
                string param = tb.Name.Contains("SetPoint") ? "SetPoint" :
                               tb.Name.Contains("Actual") ? "Actual" :
                               tb.Name.Contains("Pitch") ? "Pitch" : "Roll";
                tb.Text = $"{param}: -";
            }

            foreach (var plot in plots)
            {
                plot.Plot.Axes.AutoScale();
                plot.Refresh();
            }

            txtTestStatus.Text = "Test Status: Idle";
            txtTestStatus.Foreground = new SolidColorBrush(System.Windows.Media.Colors.LightGray);

            // Re-enable button after animation & clear
            if (sender is Button btn)
            {
                btn.IsEnabled = true;
            }
        }

        private async void PlotFromExcel_Click(object sender, RoutedEventArgs e)
        {
            ClearAll_Click(null, null);

            var dialog = new OpenFileDialog
            {
                Title = "Select Excel File",
                Filter = "Excel Files (*.xlsx)|*.xlsx",
                Multiselect = false
            };

            if (dialog.ShowDialog() != true) return;

            statusText.Text = "Loading Excel file... (will plot gradually)";
            Mouse.OverrideCursor = Cursors.Wait;

            try
            {
                await LoadExcelToQueueGraduallyAsync(dialog.FileName);

                // Update plot titles with found serial numbers
                // به‌روزرسانی عنوان پلات‌ها بر اساس سریال واقعی موجود در داده‌های لود شده
                for (int i = 0; i < deviceCount; i++)
                {
                    int testId = i + 1;

                    // آخرین داده این دستگاه رو پیدا کن
                    var lastData = GlobalData.DBList
                        .Where(d => d.TestID == testId)
                        .OrderByDescending(d => d.Time)
                        .FirstOrDefault();

                    string serialToUse = deviceSerialNumbers[i]; // fallback به پیش‌فرض (SN-DEV00x)

                    if (lastData != null && !string.IsNullOrWhiteSpace(lastData.SerialNumber))
                    {
                        serialToUse = lastData.SerialNumber.Trim();
                    }

                    // آپدیت آرایه اصلی و عنوان پلات
                    deviceSerialNumbers[i] = serialToUse;
                    plots[i].Plot.Axes.Title.Label.Text = $"Device {i + 1} - {serialToUse}";
                    plots[i].Plot.Axes.AutoScale(); // اختیاری اما خوبه
                    plots[i].Refresh();
                }

                CurrentDeviceSerials = (string[])deviceSerialNumbers.Clone();

                statusText.Text = $"Successfully loaded - {GlobalData.DBList.Count:N0} points (plotting gradually)";

                if (isPaused)
                {
                    PauseResume_Click(null, null);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading Excel file:\n{ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                statusText.Text = "Loading failed";
            }
            finally
            {
                Mouse.OverrideCursor = null;
            }
        }

        private async Task LoadExcelToQueueGraduallyAsync(string filePath)
        {
            const int CHUNK_SIZE = 3000;
            const int DELAY_MS_BETWEEN_CHUNKS = 20;

            await Task.Run(() =>
            {
                using (var workbook = new XLWorkbook(filePath))
                {
                    var allData = new List<DeviceData>();
                    int totalProcessed = 0;

                    foreach (var worksheet in workbook.Worksheets)
                    {
                        if (worksheet.Name == "Summary" || (!worksheet.Name.StartsWith("Device_") && !worksheet.Name.StartsWith("Dev_")))
                            continue;

                        string deviceKey = worksheet.Name;
                        int totalRowsThisSheet = worksheet.LastRowUsed()?.RowNumber() ?? 0;

                        int currentRow = 2;

                        while (currentRow <= totalRowsThisSheet)
                        {
                            int rowsThisChunk = Math.Min(CHUNK_SIZE, totalRowsThisSheet - currentRow + 1);

                            for (int i = 0; i < rowsThisChunk; i++)
                            {
                                int row = currentRow + i;

                                if (worksheet.Cell(row, 1).IsEmpty())
                                    break;

                                if (!worksheet.Cell(row, 1).TryGetValue(out int testId))
                                    continue;

                                string serialNumber = string.Empty;
                                int time = 0;
                                double setPoint = 0, actual = 0, pitch = 0, roll = 0;

                                worksheet.Cell(row, 2).TryGetValue(out serialNumber);
                                worksheet.Cell(row, 3).TryGetValue(out time);
                                worksheet.Cell(row, 4).TryGetValue(out setPoint);
                                worksheet.Cell(row, 5).TryGetValue(out actual);
                                worksheet.Cell(row, 6).TryGetValue(out pitch);
                                worksheet.Cell(row, 7).TryGetValue(out roll);

                                var data = new DeviceData
                                {
                                    TestID = testId,
                                    Time = time,
                                    SetPoint = setPoint,
                                    Actual = actual,
                                    Pitch = pitch,
                                    Roll = roll,
                                    SerialNumber = string.IsNullOrWhiteSpace(serialNumber)
                                        ? deviceKey.Replace("Device_", "SN-DEV")
                                        : serialNumber.Trim()
                                };

                                if (string.IsNullOrWhiteSpace(data.SerialNumber) ||
                                    data.SerialNumber.StartsWith("SN-DEV") ||
                                    data.SerialNumber.StartsWith("SN-UNKNOWN"))
                                {
                                    int devIndex = testId - 1;
                                    if (devIndex >= 0 && devIndex < CurrentDeviceSerials.Length)
                                        data.SerialNumber = CurrentDeviceSerials[devIndex];
                                }

                                allData.Add(data);
                                totalProcessed++;
                            }

                            currentRow += rowsThisChunk;

                            Dispatcher.Invoke(() =>
                            {
                                statusText.Text = $"Loading sheets... {totalProcessed} points processed";
                            });

                            Thread.Sleep(DELAY_MS_BETWEEN_CHUNKS);
                        }
                    }

                    // Final sort: by TestID then Time
                    allData.Sort((a, b) =>
                    {
                        int testCompare = a.TestID.CompareTo(b.TestID);
                        return testCompare != 0 ? testCompare : a.Time.CompareTo(b.Time);
                    });

                    // Enqueue all sorted data
                    foreach (var data in allData)
                    {
                        GlobalData.DiagramQueue.Enqueue(data);
                    }

                    Dispatcher.Invoke(() =>
                    {
                        statusText.Text = $"All sheets loaded - {allData.Count:N0} points (plotting gradually)";
                    });
                }
            });
        }

        private void PlotCheckbox_Changed(object sender, RoutedEventArgs e)
        {
            if (loggerLists == null || plots == null) return;
            if (!(sender is CheckBox cb)) return;
            if (!(cb.Tag is string tagStr)) return;

            var parts = tagStr.Split(';');
            if (parts.Length != 2) return;

            if (!int.TryParse(parts[0], out int plotIndex) || !int.TryParse(parts[1], out int loggerIndex))
                return;

            if (plotIndex < 0 || plotIndex >= loggerLists.Length) return;
            if (loggerIndex < 0 || loggerIndex >= loggerLists[plotIndex].Count) return;

            loggerLists[plotIndex][loggerIndex].IsVisible = cb.IsChecked == true;
            plots[plotIndex].Refresh();
        }

        private void TestLoad_Click(object sender, RoutedEventArgs e)
        {
            if (serialPortReader == null || !serialPortReader.IsConnected)
            {
                MessageBox.Show("Please connect to a serial port first.", "Not Connected",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            UpdateSerialNumbersFromUI();
            ClearAll_Click(null, null);
            serialPortReader.SendCommand("S");
            txtTestStatus.Text = "Test Status: Running (Load)";
            txtTestStatus.Foreground = new SolidColorBrush(System.Windows.Media.Colors.Yellow);
            statusText.Text = "Test Load Started | Waiting for data...";

            if (isPaused)
            {
                PauseResume_Click(null, null);
            }
        }

        private void TestNoLoad_Click(object sender, RoutedEventArgs e)
        {
            if (serialPortReader == null || !serialPortReader.IsConnected)
            {
                MessageBox.Show("Please connect to a serial port first.", "Not Connected",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            UpdateSerialNumbersFromUI();
            ClearAll_Click(null, null);
            serialPortReader.SendCommand("TEST:NOLOAD\n");
            txtTestStatus.Text = "Test Status: Running (No Load)";
            txtTestStatus.Foreground = new SolidColorBrush(System.Windows.Media.Colors.Yellow);
            statusText.Text = "Test NoLoad Started | Waiting for data...";

            if (isPaused)
            {
                PauseResume_Click(null, null);
            }
        }

        private void SaveToExcel_Click(object sender, RoutedEventArgs e)
        {
            if (GlobalData.DBList.Count == 0)
            {
                MessageBox.Show("No data to save.", "Empty Data",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            UpdateSerialNumbersInDBList();
            excelOperation.SaveToExcel();
            GlobalData.DBList.Clear();
            ClearAll_Click(null, null);
            txtTestStatus.Text = "Test Status: Idle";
            txtTestStatus.Foreground = new SolidColorBrush(System.Windows.Media.Colors.LightGray);
            statusText.Text = "Data saved and cleared | Ready for new test";
        }

        private void SaveToPdf_Click(object sender, RoutedEventArgs e)
        {
            // TODO: Implement PDF export if needed
            MessageBox.Show("PDF export not implemented yet.", "Info", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void UpdateSerialNumbersInDBList()
        {
            UpdateSerialNumbersFromUI();
            foreach (var data in GlobalData.DBList)
            {
                int deviceIndex = data.TestID - 1;
                if (deviceIndex >= 0 && deviceIndex < deviceCount)
                {
                    data.SerialNumber = deviceSerialNumbers[deviceIndex];
                }
            }
        }

        private void UpdateSerialNumbersFromUI()
        {
            deviceSerialNumbers[0] = string.IsNullOrWhiteSpace(txtSerial1.Text) ? "SN-DEV001" : txtSerial1.Text.Trim();
            deviceSerialNumbers[1] = string.IsNullOrWhiteSpace(txtSerial2.Text) ? "SN-DEV002" : txtSerial2.Text.Trim();
            deviceSerialNumbers[2] = string.IsNullOrWhiteSpace(txtSerial3.Text) ? "SN-DEV003" : txtSerial3.Text.Trim();
            deviceSerialNumbers[3] = string.IsNullOrWhiteSpace(txtSerial4.Text) ? "SN-DEV004" : txtSerial4.Text.Trim();

            for (int i = 0; i < deviceCount; i++)
            {
                plots[i].Plot.Axes.Title.Label.Text = $"Device {i + 1} - {deviceSerialNumbers[i]}";
                plots[i].Refresh();
            }

            CurrentDeviceSerials = (string[])deviceSerialNumbers.Clone();
        }

        private void SerialTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (!(sender is TextBox tb)) return;

            int index = -1;
            if (tb.Name == "txtSerial1") index = 0;
            else if (tb.Name == "txtSerial2") index = 1;
            else if (tb.Name == "txtSerial3") index = 2;
            else if (tb.Name == "txtSerial4") index = 3;

            if (index >= 0 && index < deviceCount)
            {
                string newSerial = string.IsNullOrWhiteSpace(tb.Text) ? $"SN-DEV{index + 1:D3}" : tb.Text.Trim();
                deviceSerialNumbers[index] = newSerial;
                plots[index].Plot.Axes.Title.Label.Text = $"Device {index + 1} - {newSerial}";
                plots[index].Refresh();
                CurrentDeviceSerials[index] = newSerial;
            }
        }

        // ────────────────────────────── Fullscreen & HD Mode Logic ──────────────────────────────

        private void BtnFullHD_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && int.TryParse(btn.Tag?.ToString(), out int deviceIndex))
            {
                if (fullScreenDevice == deviceIndex)
                {
                    ResetToNormalView();
                }
                else
                {
                    // FullHD mode: Full screen with window resize to 1920x1080
                    this.WindowState = WindowState.Normal;
                    this.Width = 1920;
                    this.Height = 1080;
                    this.Left = (SystemParameters.PrimaryScreenWidth - this.Width) / 2;
                    this.Top = (SystemParameters.PrimaryScreenHeight - this.Height) / 2;

                    MakeFullScreen(deviceIndex);
                }
            }
        }

        private void BtnHD_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && int.TryParse(btn.Tag?.ToString(), out int deviceIndex))
            {
                if (hdModeDevice == deviceIndex)
                {
                    ResetToNormalView();
                }
                else
                {
                    MakeHDMode(deviceIndex);
                }
            }
        }

        private void MakeFullScreen(int deviceIndex)
        {
            fullScreenDevice = deviceIndex;
            hdModeDevice = null; // Ensure HD mode is not active

            leftPanelColumn.Width = new GridLength(0);

            plotsGrid.Children.Clear();
            plotsGrid.RowDefinitions.Clear();
            plotsGrid.ColumnDefinitions.Clear();

            plotsGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            plotsGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            Grid selectedGrid = GetPlotGrid(deviceIndex);
            if (selectedGrid != null)
            {
                if (selectedGrid.Parent is Panel parentPanel)
                {
                    parentPanel.Children.Remove(selectedGrid);
                }

                selectedGrid.Margin = new Thickness(0);
                plotsGrid.Children.Add(selectedGrid);
                Grid.SetRow(selectedGrid, 0);
                Grid.SetColumn(selectedGrid, 0);
            }

            GetFullHDButton(deviceIndex).Content = "Back";

            plots[deviceIndex].Plot.Axes.AutoScale();
            plots[deviceIndex].Refresh();

            statusText.Text = $"FullHD mode - Device {deviceIndex + 1}";
        }

        private void MakeHDMode(int deviceIndex)
        {
            // Reset first to avoid conflicts
            ResetToNormalView();

            hdModeDevice = deviceIndex;
            fullScreenDevice = null;

            // Hide left panel
            leftPanelColumn.Width = new GridLength(0);

            // Convert plotsGrid to single large cell
            plotsGrid.Children.Clear();
            plotsGrid.RowDefinitions.Clear();
            plotsGrid.ColumnDefinitions.Clear();

            plotsGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            plotsGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            Grid selectedGrid = GetPlotGrid(deviceIndex);
            if (selectedGrid != null)
            {
                if (selectedGrid.Parent is Panel parentPanel)
                {
                    parentPanel.Children.Remove(selectedGrid);
                }

                selectedGrid.Margin = new Thickness(0);
                plotsGrid.Children.Add(selectedGrid);
                Grid.SetRow(selectedGrid, 0);
                Grid.SetColumn(selectedGrid, 0);
            }

            GetHDButton(deviceIndex).Content = "Back";

            plots[deviceIndex].Plot.Axes.AutoScale();
            plots[deviceIndex].Refresh();

            statusText.Text = $"HD mode - Device {deviceIndex + 1} (full grid size)";
        }

        private void ResetToNormalView()
        {
            fullScreenDevice = null;
            hdModeDevice = null;

            leftPanelColumn.Width = new GridLength(380);

            plotsGrid.Children.Clear();
            plotsGrid.RowDefinitions.Clear();
            plotsGrid.ColumnDefinitions.Clear();

            plotsGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            plotsGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            plotsGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            plotsGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            for (int i = 0; i < deviceCount; i++)
            {
                Grid grid = GetPlotGrid(i);
                if (grid != null)
                {
                    grid.Margin = new Thickness(10);
                    plotsGrid.Children.Add(grid);
                    Grid.SetRow(grid, i / 2);
                    Grid.SetColumn(grid, i % 2);
                }

                GetFullHDButton(i).Content = "FullHD";
                GetHDButton(i).Content = "HD";
            }

            foreach (var plot in plots)
            {
                plot.Plot.Axes.AutoScale();
                plot.Refresh();
            }

            statusText.Text = "Returned to normal view | Ready";
        }

        private Grid GetPlotGrid(int index)
        {
            Grid result = null;
            switch (index)
            {
                case 0:
                    result = grid0;
                    break;
                case 1:
                    result = grid1;
                    break;
                case 2:
                    result = grid2;
                    break;
                case 3:
                    result = grid3;
                    break;
            }
            return result;
        }

        private Button GetFullHDButton(int index)
        {
            Button result = null;
            switch (index)
            {
                case 0:
                    result = btnFullHD0;
                    break;
                case 1:
                    result = btnFullHD1;
                    break;
                case 2:
                    result = btnFullHD2;
                    break;
                case 3:
                    result = btnFullHD3;
                    break;
            }
            return result;
        }

        private Button GetHDButton(int index)
        {
            Button result = null;
            switch (index)
            {
                case 0:
                    result = btnHD0;
                    break;
                case 1:
                    result = btnHD1;
                    break;
                case 2:
                    result = btnHD2;
                    break;
                case 3:
                    result = btnHD3;
                    break;
            }
            return result;
        }
    }
}