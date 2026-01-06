using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.IO.Ports;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Threading;
using Microsoft.Win32;
using ScottPlot;
using ScottPlot.Plottables;
using ScottPlot.WPF;
using ClosedXML.Excel;
using System.Threading;
using System.Text;
using System.IO;

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
        private int lastProcessedIndex = 0;

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
            if (isPaused) return;

            int start = lastProcessedIndex;
            int count = GlobalData.DBList.Count - start;
            if (count <= 0) return;

            bool updated = false;
            var updatedPlots = new HashSet<int>();
            double currentMaxTime = lastTimeIndex;

            for (int k = 0; k < count; k++)
            {
                var data = GlobalData.DBList[start + k];
                int deviceIndex = data.TestID - 1;
                if (deviceIndex < 0 || deviceIndex >= deviceCount) continue;

                // Update serial number if necessary
                if (string.IsNullOrWhiteSpace(data.SerialNumber) || data.SerialNumber.StartsWith("SN-DEV"))
                {
                    data.SerialNumber = deviceSerialNumbers[deviceIndex];
                }

                var loggerList = loggerLists[deviceIndex];

                // Add data to plot (no removal – all data stays)
                loggerList[0].Add(data.Time, data.SetPointDeg);
                loggerList[1].Add(data.Time, data.ActualDeg);
                loggerList[2].Add(data.Time, data.PitchDeg);
                loggerList[3].Add(data.Time, data.RollDeg);

                // Update live values
                int baseIndex = deviceIndex * 4;
                liveValueTexts[baseIndex + 0].Text = $"SetPoint: {data.SetPointDeg:F3}";
                liveValueTexts[baseIndex + 1].Text = $"Actual: {data.ActualDeg:F3}";
                liveValueTexts[baseIndex + 2].Text = $"Pitch: {data.PitchDeg:F3}";
                liveValueTexts[baseIndex + 3].Text = $"Roll: {data.RollDeg:F3}";

                updated = true;
                updatedPlots.Add(deviceIndex);

                if (data.Time > currentMaxTime) currentMaxTime = data.Time;
            }

            lastProcessedIndex += count;
            totalPoints += count;

            livePointsText.Text = totalPoints.ToString("N0");
            liveLastUpdate.Text = currentMaxTime.ToString("F0");
            lastTimeIndex = currentMaxTime;

            // Update status bar
            statusText.Text = $"Running | Points: {totalPoints:N0} | Plotted: {lastProcessedIndex:N0}/{GlobalData.DBList.Count:N0}";

            if (updated)
            {
                foreach (int index in updatedPlots)
                {
                    plots[index].Plot.Axes.AutoScale();
                    plots[index].Refresh();
                }
            }

            // Update device titles if serial changed
            bool titlesUpdated = false;
            for (int i = 0; i < deviceCount; i++)
            {
                int testId = i + 1;
                var anyData = GlobalData.DBList.FirstOrDefault(d => d.TestID == testId && !string.IsNullOrWhiteSpace(d.SerialNumber));
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
            if (lastProcessedIndex >= GlobalData.DBList.Count)
            {
                statusText.Text = $"Plot finished | Total points: {totalPoints:N0}";
            }
            else
            {
                statusText.Text = $"Plotting... {lastProcessedIndex:N0}/{GlobalData.DBList.Count:N0} points";
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
            // Disable button during animation & clear
            if (sender is Button clearButton)
            {
                clearButton.IsEnabled = false;

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

                await Task.Delay(300);
            }

            // Clear everything
            foreach (var loggerList in loggerLists)
            {
                foreach (var logger in loggerList)
                {
                    logger.Data.Clear();
                }
            }

            totalPoints = 0;
            lastTimeIndex = -1;
            lastProcessedIndex = 0;

            GlobalData.DBList.Clear();

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

            statusText.Text = "Loading Excel file...";
            Mouse.OverrideCursor = Cursors.Wait;

            try
            {
                var allData = new List<DeviceData>();

                using (var workbook = new XLWorkbook(dialog.FileName))
                {
                    int totalProcessed = 0;
                    int totalRowsExpected = 0;

                    foreach (var worksheet in workbook.Worksheets)
                    {
                        if (worksheet.Name == "Summary" ||
                            (!worksheet.Name.StartsWith("Device_") && !worksheet.Name.StartsWith("Dev_")))
                            continue;

                        totalRowsExpected += (worksheet.LastRowUsed()?.RowNumber() ?? 0) - 1;
                    }

                    foreach (var worksheet in workbook.Worksheets)
                    {
                        if (worksheet.Name == "Summary" ||
                            (!worksheet.Name.StartsWith("Device_") && !worksheet.Name.StartsWith("Dev_")))
                            continue;

                        int columnCount = worksheet.Row(1).CellsUsed().Count();
                        int totalRows = worksheet.LastRowUsed()?.RowNumber() ?? 0;

                        for (int row = 2; row <= totalRows; row++)
                        {
                            if (worksheet.Cell(row, 1).IsEmpty()) continue;
                            if (!worksheet.Cell(row, 1).TryGetValue(out int testId)) continue;

                            string serialNumber = string.Empty;
                            int time = 0;
                            double setPointRaw = 0, actualRaw = 0, pitchRaw = 0, rollRaw = 0;
                            double setPointDeg = 0, actualDeg = 0, pitchDeg = 0, rollDeg = 0;

                            worksheet.Cell(row, 2).TryGetValue(out serialNumber);
                            worksheet.Cell(row, 3).TryGetValue(out time);

                            if (columnCount >= 11)
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
                                SerialNumber = string.IsNullOrWhiteSpace(serialNumber) ? "" : serialNumber.Trim()
                            };

                            int devIndex = testId - 1;
                            if (string.IsNullOrWhiteSpace(data.SerialNumber) ||
                                data.SerialNumber.StartsWith("SN-DEV") ||
                                data.SerialNumber.StartsWith("SN-UNKNOWN"))
                            {
                                if (devIndex >= 0 && devIndex < CurrentDeviceSerials.Length)
                                {
                                    data.SerialNumber = CurrentDeviceSerials[devIndex];
                                }
                            }

                            allData.Add(data);
                            totalProcessed++;

                            statusText.Text = $"Loading from Excel... {totalProcessed:N0}/{totalRowsExpected:N0} points loaded";
                        }
                    }
                }

                allData.Sort((a, b) =>
                {
                    int timeCmp = a.Time.CompareTo(b.Time);
                    return timeCmp != 0 ? timeCmp : a.TestID.CompareTo(b.TestID);
                });

                GlobalData.DBList.AddRange(allData);

                lastProcessedIndex = 0;
                totalPoints = 0;
                lastTimeIndex = -1;

                statusText.Text = $"Excel loaded ({allData.Count:N0} points) – Plotting gradually...";

                if (isPaused)
                {
                    PauseResume_Click(null, null); // Resume if paused
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
                        if (worksheet.Name == "Summary" ||
                            (!worksheet.Name.StartsWith("Device_") && !worksheet.Name.StartsWith("Dev_")))
                            continue;

                        int columnCount = worksheet.Row(1).CellsUsed().Count();

                        int totalRowsThisSheet = worksheet.LastRowUsed()?.RowNumber() ?? 0;
                        int currentRow = 2;

                        while (currentRow <= totalRowsThisSheet)
                        {
                            int rowsThisChunk = Math.Min(CHUNK_SIZE, totalRowsThisSheet - currentRow + 1);

                            for (int i = 0; i < rowsThisChunk; i++)
                            {
                                int row = currentRow + i;
                                if (worksheet.Cell(row, 1).IsEmpty()) break;

                                if (!worksheet.Cell(row, 1).TryGetValue(out int testId)) continue;

                                string serialNumber = string.Empty;
                                int time = 0;

                                double setPointRaw = 0, actualRaw = 0, pitchRaw = 0, rollRaw = 0;
                                double setPointDeg = 0, actualDeg = 0, pitchDeg = 0, rollDeg = 0;

                                worksheet.Cell(row, 2).TryGetValue(out serialNumber);
                                worksheet.Cell(row, 3).TryGetValue(out time);

                                if (columnCount >= 11)
                                {
                                    // New format: Raw + Deg
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
                                    // Old format: only Deg
                                    worksheet.Cell(row, 4).TryGetValue(out setPointDeg);
                                    worksheet.Cell(row, 5).TryGetValue(out actualDeg);
                                    worksheet.Cell(row, 6).TryGetValue(out pitchDeg);
                                    worksheet.Cell(row, 7).TryGetValue(out rollDeg);

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
                                    SerialNumber = string.IsNullOrWhiteSpace(serialNumber) ? "" : serialNumber.Trim()
                                };

                                // Fallback for serial number
                                int devIndex = testId - 1;
                                if (string.IsNullOrWhiteSpace(data.SerialNumber) ||
                                    data.SerialNumber.StartsWith("SN-DEV") ||
                                    data.SerialNumber.StartsWith("SN-UNKNOWN"))
                                {
                                    if (devIndex >= 0 && devIndex < CurrentDeviceSerials.Length)
                                    {
                                        data.SerialNumber = CurrentDeviceSerials[devIndex];
                                    }
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

                    // Global sort: primarily by Time, secondarily by TestID
                    allData.Sort((a, b) =>
                    {
                        int timeCmp = a.Time.CompareTo(b.Time);
                        return timeCmp != 0 ? timeCmp : a.TestID.CompareTo(b.TestID);
                    });

                    // Add all sorted data to DBList on UI thread
                    Dispatcher.Invoke(() =>
                    {
                        GlobalData.DBList.AddRange(allData);
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
            if (!int.TryParse(parts[0], out int plotIndex) || !int.TryParse(parts[1], out int loggerIndex)) return;

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

        // Fullscreen & HD Mode Logic
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
            hdModeDevice = null;

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
            ResetToNormalView();

            hdModeDevice = deviceIndex;
            fullScreenDevice = null;

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
            switch (index)
            {
                case 0: return grid0;
                case 1: return grid1;
                case 2: return grid2;
                case 3: return grid3;
                default: return null;
            }
        }

        private Button GetFullHDButton(int index)
        {
            switch (index)
            {
                case 0: return btnFullHD0;
                case 1: return btnFullHD1;
                case 2: return btnFullHD2;
                case 3: return btnFullHD3;
                default: return null;
            }
        }

        private Button GetHDButton(int index)
        {
            switch (index)
            {
                case 0: return btnHD0;
                case 1: return btnHD1;
                case 2: return btnHD2;
                case 3: return btnHD3;
                default: return null;
            }
        }
        private void BtnAnalyze_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && int.TryParse(button.Tag?.ToString(), out int deviceIndex))
            {
                int testId = deviceIndex + 1;

                var deviceData = GlobalData.DBList.Where(d => d.TestID == testId).OrderBy(d => d.Time).ToList();

                if (deviceData.Count < 100)
                {
                    MessageBox.Show("Not enough data for analysis.", "Insufficient Data", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var dialog = new AnalysisTypeDialog(testId);
                if (dialog.ShowDialog() == true)
                {
                    if (dialog.DoNoLoad)
                    {

                        string noLoadReport = NoLoad.Analyze(deviceData, testId);


                        var previewWindow = new ReportPreviewWindow(noLoadReport, testId);
                        previewWindow.ShowDialog();
                    }
                }
            }
        }
    }
}