using System;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using Microsoft.Win32;
using iText.Kernel.Pdf;
using iText.Kernel.Geom;
using iText.Layout;
using iText.Layout.Element;
using iText.Layout.Borders;
using iText.Kernel.Colors;
using iText.Kernel.Font;
using iText.IO.Font.Constants;
using iText.Layout.Properties;

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

            if (saveDialog.ShowDialog() != true) return;

            try
            {
                using (var writer = new PdfWriter(saveDialog.FileName))
                using (var pdf = new PdfDocument(writer))
                using (var document = new Document(pdf, PageSize.A4))
                {
                    document.SetMargins(70, 50, 80, 50);

                    // Professional color scheme – modern and clean
                    var primaryBlue = new DeviceRgb(0, 70, 140);           // Deep blue for titles/headers
                    var lightBlueBg = new DeviceRgb(230, 240, 255);        // Light blue background for sections
                    var alternateRow = new DeviceRgb(245, 250, 255);       // Alternating rows in tables
                    var passGreen = new DeviceRgb(0, 120, 0);              // Dark green text for PASS
                    var passCellBg = new DeviceRgb(220, 255, 220);         // Light green background for PASS cells

                    // Fonts
                    var titleFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);
                    var sectionFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);
                    var labelFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);
                    var normalFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA);
                    var headerFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);

                    var lines = ReportText.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);

                    Table currentTable = null;
                    int currentSection = 0;
                    int rowIndex = 0;

                    // Optimized column widths
                    var stepColumnWidths = new float[] { 13f, 13f, 15f, 15f, 15f, 12f, 17f };
                    var sensColumnWidths = new float[] { 25f, 25f, 25f, 25f };

                    var stepHeaders = new[] { "From", "To", "DeadT(ms)", "RiseT(ms)", "SetT(ms)", "OS(%)", "SSE(deg)" };
                    var sensHeaders = new[] { "Target", "Actual", "Error", "Status" };

                    foreach (var rawLine in lines)
                    {
                        var line = rawLine.TrimEnd();
                        if (string.IsNullOrWhiteSpace(line)) continue;

                        // Completely ignore all separator lines (----, ====, etc.) – no gray bars added
                        if (line.Length > 15 && line.All(c => c == '-' || c == '=' || c == '_')) continue;

                        // Main title – added only once, large and blue
                        if (line.Contains("NO-LOAD CONTROL SYSTEM ANALYSIS REPORT"))
                        {
                            document.Add(new Paragraph("NO-LOAD CONTROL SYSTEM ANALYSIS REPORT")
                                .SetFont(titleFont)
                                .SetFontSize(26)
                                .SetFontColor(primaryBlue)
                                .SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER)
                                .SetMarginBottom(50));
                            continue;
                        }

                        // Section titles with light blue background
                        if (line.StartsWith("1.") || line.StartsWith("2.") || line.StartsWith("3.") || line.StartsWith("4."))
                        {
                            if (currentTable != null)
                            {
                                document.Add(currentTable.SetMarginBottom(40));
                                currentTable = null;
                                rowIndex = 0;
                            }

                            currentSection = int.Parse(line.Substring(0, 1));

                            document.Add(new Paragraph(line)
                                .SetFont(sectionFont)
                                .SetFontSize(16)
                                .SetFontColor(primaryBlue)
                                .SetPadding(12)
                                .SetBackgroundColor(lightBlueBg)
                                .SetMarginTop(40)
                                .SetMarginBottom(20));
                            continue;
                        }

                        // Key-value information lines (centered, blue labels)
                        if (line.Contains(":"))
                        {
                            int colon = line.IndexOf(':');
                            string label = line.Substring(0, colon + 1).Trim();
                            string value = line.Substring(colon + 1).Trim();

                            var p = new Paragraph()
                                .SetFontSize(12)
                                .SetFixedLeading(22)
                                .SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER)
                                .SetMarginBottom(10);

                            p.Add(new Text(label + " ").SetFont(labelFont).SetFontColor(primaryBlue));
                            p.Add(new Text(value).SetFont(normalFont));

                            document.Add(p);
                            continue;
                        }

                        // Start single large Step Response table
                        if (currentSection == 3 && line.Contains("From") && line.Contains("To") && line.Contains("DeadT"))
                        {
                            if (currentTable != null) document.Add(currentTable.SetMarginBottom(40));

                            currentTable = new Table(UnitValue.CreatePercentArray(stepColumnWidths))
                                .UseAllAvailableWidth()
                                .SetMarginTop(10);

                            foreach (var h in stepHeaders)
                            {
                                currentTable.AddHeaderCell(new Cell()
                                    .Add(new Paragraph(h).SetFont(headerFont).SetFontSize(11).SetFontColor(ColorConstants.WHITE))
                                    .SetBackgroundColor(primaryBlue)
                                    .SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER)
                                    .SetPadding(10));
                            }

                            rowIndex = 0;
                            continue;
                        }

                        // Start Sensitivity table
                        if (currentSection == 4 && line.Contains("Target") && line.Contains("Actual") && line.Contains("Status"))
                        {
                            if (currentTable != null) document.Add(currentTable.SetMarginBottom(40));

                            currentTable = new Table(UnitValue.CreatePercentArray(sensColumnWidths))
                                .UseAllAvailableWidth()
                                .SetMarginTop(10);

                            foreach (var h in sensHeaders)
                            {
                                currentTable.AddHeaderCell(new Cell()
                                    .Add(new Paragraph(h).SetFont(headerFont).SetFontSize(11).SetFontColor(ColorConstants.WHITE))
                                    .SetBackgroundColor(primaryBlue)
                                    .SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER)
                                    .SetPadding(10));
                            }

                            rowIndex = 0;
                            continue;
                        }

                        // Add rows to current table (single unified table – no splits)
                        if (currentTable != null)
                        {
                            var cells = line.Split(new char[0], StringSplitOptions.RemoveEmptyEntries)
                                            .Select(c => c.Trim())
                                            .ToArray();

                            if (cells.Length == currentTable.GetNumberOfColumns())
                            {
                                Color rowBg = (rowIndex % 2 == 0) ? ColorConstants.WHITE : alternateRow;

                                for (int i = 0; i < cells.Length; i++)
                                {
                                    string cellText = cells[i];
                                    var cellPara = new Paragraph(cellText)
                                        .SetFont(normalFont)
                                        .SetFontSize(11);

                                    var cell = new Cell()
                                        .Add(cellPara)
                                        .SetBackgroundColor(rowBg)
                                        .SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER)
                                        .SetPadding(9)
                                        .SetBorder(new SolidBorder(new DeviceGray(0.7f), 0.5f));

                                    if (currentSection == 4 && i == 3 && cellText == "PASS")
                                    {
                                        cellPara.SetFontColor(passGreen)
                                                .SetFont(headerFont); // bold
                                        cell.SetBackgroundColor(passCellBg);
                                    }

                                    currentTable.AddCell(cell);
                                }

                                rowIndex++;
                            }
                            continue;
                        }

                        // Fallback
                        document.Add(new Paragraph(line).SetFont(normalFont).SetFontSize(12).SetMarginBottom(8));
                    }

                    if (currentTable != null) document.Add(currentTable.SetMarginBottom(40));
                }

                MessageBox.Show("PDF report saved successfully with a highly professional, modern, and beautiful design!\n" +
                                "• Single unified tables (no splits or gray bars)\n" +
                                "• Highlighted PASS cells (green + bold)\n" +
                                "• Alternating row colors\n" +
                                "• Blue theme with clean spacing\n" +
                                "The old gray separator bars are completely gone.", "Saved",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error saving PDF:\n{ex.Message}", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}