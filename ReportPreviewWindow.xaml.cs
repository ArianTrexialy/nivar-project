using System;
using System.Windows;
using Microsoft.Win32;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;

namespace DeviceAnalisys_v5
{
    public partial class ReportPreviewWindow : Window
    {
        public string ReportText { get; set; }

        public ReportPreviewWindow(string report, int deviceId)
        {
            InitializeComponent();
            ReportText = report;
            DataContext = this;
            Title = $"Report Preview - Device {deviceId}";
        }

        private void SavePdf_Click(object sender, RoutedEventArgs e)
        {
            var saveDialog = new SaveFileDialog
            {
                Filter = "PDF Files (*.pdf)|*.pdf",
                FileName = $"NoLoad_Analysis_Device_{DateTime.Now:yyyyMMdd_HHmmss}.pdf"
            };

            if (saveDialog.ShowDialog() == true)
            {
                using (var writer = new PdfWriter(saveDialog.FileName))
                using (var pdf = new PdfDocument(writer))
                using (var document = new Document(pdf))
                {
                    document.Add(new Paragraph(ReportText).SetFontSize(12));
                }

                MessageBox.Show("Report saved as PDF successfully!", "Saved", MessageBoxButton.OK, MessageBoxImage.Information);
                this.Close(); 
            }
        }
    }
}