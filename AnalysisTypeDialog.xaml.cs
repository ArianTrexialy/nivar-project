using System.Windows;

namespace DeviceAnalisys_v5
{
    public partial class AnalysisTypeDialog : Window
    {
        public string DeviceTitle { get; }
        public bool DoNoLoad => rbNoLoad.IsChecked == true || rbBoth.IsChecked == true;
        public bool DoLoad => rbLoad.IsChecked == true || rbBoth.IsChecked == true;

        public AnalysisTypeDialog(int deviceId)
        {
            InitializeComponent();
            DeviceTitle = $"Analyze Device {deviceId}";
            DataContext = this;
        }

        private void Analyze_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
            Close();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}