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

namespace DeviceAnalisys_v5
{
    //this is my 4th edit
    public partial class MainWindow : Window
    {
        private ExcelOperation excelOperation = new ExcelOperation();
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
        private bool _isPaused = false;
        private int _totalPoints = 0;
        private double _lastTimeIndex = -1;
        private int? fullScreenDevice = null;
        private ColumnDefinition leftPanelColumn; // این خط جدیده — اضافه کن

        public MainWindow()
        {
            InitializeComponent();
            PlotInitial();
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
                if (_isPaused)
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

        private void PlotInitial()
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
                plt.Legend.Location = Alignment.UpperRight;
                plt.Legend.FontSize = 12;
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

            plotTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(2000) };
            plotTimer.Tick += PlotTimer_Tick;
            plotTimer.Start();

            cmbFps.SelectionChanged -= cmbFps_SelectionChanged;
            cmbFps.SelectionChanged += cmbFps_SelectionChanged;
        }

        private void PlotTimer_Tick(object sender, EventArgs e)
        {
            if (_isPaused) return;

            const int MAX_PER_TICK = 1500;
            const int MAX_POINTS_PER_PLOT = 35000;
            bool updated = false;
            HashSet<int> updatedPlots = new HashSet<int>();
            double currentMaxTime = _lastTimeIndex;
            int processed = 0;

            while (processed < MAX_PER_TICK && GlobalData.DiagramQueue.TryDequeue(out DeviceData data))
            {
                int deviceIndex = data.TestID - 1;
                if (deviceIndex < 0 || deviceIndex >= deviceCount) continue;

                if (string.IsNullOrWhiteSpace(data.SerialNumber) ||
                    data.SerialNumber.StartsWith("SN-DEV") ||
                    data.SerialNumber.StartsWith("SN-UNKNOWN") ||
                    data.SerialNumber == $"SN-DEV{data.TestID:D3}")
                {
                    data.SerialNumber = deviceSerialNumbers[deviceIndex];
                }

                GlobalData.DBList.Add(data);
                var loggerList = loggerLists[deviceIndex];

                loggerList[0].Add(data.Time, data.SetPoint);
                loggerList[1].Add(data.Time, data.Actual);
                loggerList[2].Add(data.Time, data.Pitch);
                loggerList[3].Add(data.Time, data.Roll);

                foreach (var logger in loggerList)
                {
                    if (logger.Data.Coordinates.Count > MAX_POINTS_PER_PLOT)
                    {
                        int excess = logger.Data.Coordinates.Count - MAX_POINTS_PER_PLOT;
                        logger.Data.Coordinates.RemoveRange(0, excess);
                    }
                }

                int baseIndex = deviceIndex * 4;
                liveValueTexts[baseIndex + 0].Text = $"SetPoint: {data.SetPoint:F3}";
                liveValueTexts[baseIndex + 1].Text = $"Actual: {data.Actual:F3}";
                liveValueTexts[baseIndex + 2].Text = $"Pitch: {data.Pitch:F3}";
                liveValueTexts[baseIndex + 3].Text = $"Roll: {data.Roll:F3}";

                _totalPoints++;
                updated = true;
                updatedPlots.Add(deviceIndex);

                if (data.Time > currentMaxTime)
                    currentMaxTime = data.Time;

                processed++;
            }

            if (updated)
            {
                foreach (int index in updatedPlots)
                {
                    plots[index].Plot.Axes.AutoScale();
                    plots[index].Refresh();
                }

                livePointsText.Text = _totalPoints.ToString();
                liveLastUpdate.Text = currentMaxTime.ToString("F0");
                _lastTimeIndex = currentMaxTime;
                statusText.Text = $"Running | Points: {_totalPoints}";
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
            _isPaused = !_isPaused;
            if (_isPaused)
            {
                plotTimer.Stop();
                ((Button)sender).Content = "Resume";
                statusText.Text = $"Paused | Points: {_totalPoints}";
            }
            else
            {
                plotTimer.Start();
                ((Button)sender).Content = "Pause";
                statusText.Text = $"Running | Points: {_totalPoints}";
            }
        }

        private void ClearAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var loggerList in loggerLists)
            {
                foreach (var logger in loggerList)
                {
                    logger.Data.Clear();
                }
            }

            _totalPoints = 0;
            statusText.Text = "Cleared | Points: 0";

            foreach (var plot in plots)
            {
                plot.Plot.Axes.AutoScale();
                plot.Refresh();
            }

            foreach (var tb in liveValueTexts)
            {
                string param = tb.Name.Contains("SetPoint") ? "SetPoint" :
                               tb.Name.Contains("Actual") ? "Actual" :
                               tb.Name.Contains("Pitch") ? "Pitch" : "Roll";
                tb.Text = $"{param}: -";
            }

            livePointsText.Text = "0";
            liveLastUpdate.Text = "-";
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

            statusText.Text = "Loading large file... Please wait";
            Mouse.OverrideCursor = Cursors.Wait;

            try
            {
                await Task.Run(() => excelOperation.LoadExcelToDiagramQueue(dialog.FileName));

                for (int i = 0; i < deviceCount; i++)
                {
                    int testId = i + 1;
                    var lastValid = GlobalData.DBList
                        .Where(d => d.TestID == testId)
                        .OrderByDescending(d => d.Time)
                        .FirstOrDefault(d => !string.IsNullOrWhiteSpace(d.SerialNumber) &&
                                             !d.SerialNumber.StartsWith("SN-DEV") &&
                                             !d.SerialNumber.StartsWith("SN-UNKNOWN"));

                    if (lastValid != null)
                        deviceSerialNumbers[i] = lastValid.SerialNumber.Trim();

                    plots[i].Plot.Axes.Title.Label.Text = $"Device {i + 1} - {deviceSerialNumbers[i]}";
                    plots[i].Plot.Axes.AutoScale();
                    plots[i].Refresh();
                }

                CurrentDeviceSerials = (string[])deviceSerialNumbers.Clone();
                statusText.Text = $"Successfully loaded - {GlobalData.DBList.Count:N0} points";
                if (_isPaused) PauseResume_Click(null, null);
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
            if (_isPaused)
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
            if (_isPaused)
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

        private void SaveToPdf_Click(object sender, RoutedEventArgs e)
        {
            // TODO: Implement PDF export if needed
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

        // ────────────────────────────── Fullscreen Logic ──────────────────────────────

        private void BtnFull_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && int.TryParse(btn.Tag?.ToString(), out int deviceIndex))
            {
                if (fullScreenDevice == deviceIndex)
                {
                    ResetToNormalView();
                }
                else
                {
                    MakeFullScreen(deviceIndex);
                }
            }
        }

        private void BtnFullHD_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && int.TryParse(btn.Tag?.ToString(), out int deviceIndex))
            {
                // اگر قبلاً در حالت fullscreen بودیم، اول برگردونیم
                if (fullScreenDevice.HasValue)
                {
                    ResetToNormalView();
                }

                // تنظیم پنجره به 1920x1080 و وسط صفحه
                this.WindowState = WindowState.Normal;
                this.Width = 1920;
                this.Height = 1080;
                this.Left = (SystemParameters.PrimaryScreenWidth - this.Width) / 2;
                this.Top = (SystemParameters.PrimaryScreenHeight - this.Height) / 2;

                // حالا fullscreen واقعی (با مخفی شدن پنل چپ)
                MakeFullScreen(deviceIndex);
            }
        }

        private void MakeFullScreen(int deviceIndex)
        {
            fullScreenDevice = deviceIndex;

            // مخفی کردن کامل پنل سمت چپ
            leftPanelColumn.Width = new GridLength(0);

            // تبدیل plotsGrid به یک سلول بزرگ
            plotsGrid.Children.Clear();
            plotsGrid.RowDefinitions.Clear();
            plotsGrid.ColumnDefinitions.Clear();

            plotsGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            plotsGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            Grid selectedGrid = GetPlotGrid(deviceIndex);
            if (selectedGrid != null)
            {
                // جدا کردن از والد قبلی
                if (selectedGrid.Parent is Panel parentPanel)
                {
                    parentPanel.Children.Remove(selectedGrid);
                }

                // حذف حاشیه برای پر کردن کامل صفحه
                selectedGrid.Margin = new Thickness(0);

                plotsGrid.Children.Add(selectedGrid);
                Grid.SetRow(selectedGrid, 0);
                Grid.SetColumn(selectedGrid, 0);
            }

            // تغییر دکمه به Back
            GetFullButton(deviceIndex).Content = "Back";

            plots[deviceIndex].Plot.Axes.AutoScale();
            plots[deviceIndex].Refresh();
        }

        private void ResetToNormalView()
        {
            fullScreenDevice = null;

            // برگرداندن پنل سمت چپ
            leftPanelColumn.Width = new GridLength(380);

            // بازسازی ساختار 2×2
            plotsGrid.Children.Clear();
            plotsGrid.RowDefinitions.Clear();
            plotsGrid.ColumnDefinitions.Clear();

            plotsGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            plotsGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            plotsGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            plotsGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            // برگرداندن همه پلات‌ها به جای اصلی
            for (int i = 0; i < deviceCount; i++)
            {
                Grid grid = GetPlotGrid(i);
                if (grid != null)
                {
                    grid.Margin = new Thickness(10); // حاشیه اصلی
                    plotsGrid.Children.Add(grid);
                    Grid.SetRow(grid, i / 2);
                    Grid.SetColumn(grid, i % 2);
                }

                GetFullButton(i).Content = "Full";
            }

            foreach (var plot in plots)
            {
                plot.Plot.Axes.AutoScale();
                plot.Refresh();
            }
        }

        private Grid GetPlotGrid(int index)
        {
            switch (index)
            {
                case 0:
                    return grid0;
                case 1:
                    return grid1;
                case 2:
                    return grid2;
                case 3:
                    return grid3;
                default:
                    return null;
            }
        }

        private Button GetFullButton(int index)
        {
            switch (index)
            {
                case 0:
                    return btnFull0;
                case 1:
                    return btnFull1;
                case 2:
                    return btnFull2;
                case 3:
                    return btnFull3;
                default:
                    return null;
            }
        }
    }
}