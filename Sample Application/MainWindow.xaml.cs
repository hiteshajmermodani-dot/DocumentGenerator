using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Media.Imaging;
using Microsoft.Win32;
using DocumentGenerator;

namespace Sample_Application
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private readonly List<string> _imagePaths = new List<string>();

        public MainWindow()
        {
            InitializeComponent();
        }

        private void BtnAddImage_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Filter = "Image Files|*.png;*.jpg;*.jpeg;*.bmp;*.gif",
                Multiselect = true
            };

            if (dialog.ShowDialog() == true)
            {
                foreach (var file in dialog.FileNames)
                {
                    _imagePaths.Add(file);
                    lstImages.Items.Add(Path.GetFileName(file));
                }
                txtImageCount.Text = $"{_imagePaths.Count} image(s) selected";
            }
        }

        private void BtnClearImages_Click(object sender, RoutedEventArgs e)
        {
            _imagePaths.Clear();
            lstImages.Items.Clear();
            txtImageCount.Text = "No images selected";
        }

        private void BtnGenerate_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new SaveFileDialog
            {
                Filter = "Word Document (*.docx)|*.docx",
                FileName = "LargeDataReport.docx"
            };

            if (dialog.ShowDialog() != true)
                return;

            try
            {
                btnGenerate.IsEnabled = false;
                txtStatus.Foreground = System.Windows.Media.Brushes.Green;
                txtStatus.Text = "Generating report...";

                GenerateLargeReport(dialog.FileName);

                txtStatus.Text = $"Report saved to: {dialog.FileName}";

                if (MessageBox.Show("Report generated successfully! Open the file?", "Done",
                    MessageBoxButton.YesNo, MessageBoxImage.Information) == MessageBoxResult.Yes)
                {
                    Process.Start(new ProcessStartInfo(dialog.FileName) { UseShellExecute = true });
                }
            }
            catch (Exception ex)
            {
                txtStatus.Foreground = System.Windows.Media.Brushes.Red;
                txtStatus.Text = "Error: " + ex.Message;
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                btnGenerate.IsEnabled = true;
            }
        }

        private void BtnGeneratePdf_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new SaveFileDialog
            {
                Filter = "PDF Document (*.pdf)|*.pdf",
                FileName = "LargeDataReport.pdf"
            };

            if (dialog.ShowDialog() != true)
                return;

            try
            {
                btnGeneratePdf.IsEnabled = false;
                txtStatus.Foreground = System.Windows.Media.Brushes.Green;
                txtStatus.Text = "Generating PDF report...";

                GenerateLargePdfReport(dialog.FileName);

                txtStatus.Text = $"PDF saved to: {dialog.FileName}";

                if (MessageBox.Show("PDF report generated successfully! Open the file?", "Done",
                    MessageBoxButton.YesNo, MessageBoxImage.Information) == MessageBoxResult.Yes)
                {
                    Process.Start(new ProcessStartInfo(dialog.FileName) { UseShellExecute = true });
                }
            }
            catch (Exception ex)
            {
                txtStatus.Foreground = System.Windows.Media.Brushes.Red;
                txtStatus.Text = "Error: " + ex.Message;
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                btnGeneratePdf.IsEnabled = true;
            }
        }

        private void GenerateLargeReport(string filePath)
        {
            var doc = new DocxDocument();
            bool autoSpacing = chkAutoSpacing.IsChecked == true;

            // Auto Spacing
            if (autoSpacing)
            {
                doc.AutoSpacing(spacingAfterPt: 8, spacingBeforePt: 2, lineSpacingMultiple: 1.15);
                doc.SetPageMargins(topPt: 72, bottomPt: 72, leftPt: 72, rightPt: 72);
                doc.SetHeadingsStartOnNewPage(maxLevel: 1);
            }

            // Header and Footer
            if (chkHeader.IsChecked == true)
            {
                doc.SetHeader(new DocxHeaderFooter()
                    .SetText("Company Quarterly Report")
                    .SetBold()
                    .SetFontSize(10)
                    .SetAlignment("center"));
            }

            if (chkFooter.IsChecked == true)
            {
                var footer = new DocxHeaderFooter()
                    .SetText("Confidential")
                    .SetAlignment("left");

                if (chkPageNumbers.IsChecked == true)
                    footer.AddPageNumber();

                doc.SetFooter(footer);
            }
            else if (chkPageNumbers.IsChecked == true)
            {
                doc.AddPageNumbers();
            }

            // ===== Title Page =====
            doc.AddHeading("Quarterly Sales Report", 1);
            doc.AddParagraph("Generated on: " + DateTime.Now.ToString("MMMM dd, yyyy"), fontSize: 12);
            doc.AddParagraph("Department: Sales & Marketing", fontSize: 12);
            doc.AddHorizontalRule();
            doc.AddParagraph("");

            // ===== Build Table of Contents =====
            var toc = new DocxTableOfContents()
                .SetTitle("Table of Contents")
                .ShowSerialNumbers(true)
                .UseLinkStyle(true)
                .AddEntry("Executive Summary", 1, "Section_ExecutiveSummary")
                .AddEntry("Regional Sales Data", 1, "Section_RegionalSales")
                .AddEntry("Product Performance", 1, "Section_ProductPerformance")
                .AddEntry("Software Licenses", 2, "Section_SoftwareLicenses")
                .AddEntry("Cloud Services", 2, "Section_CloudServices")
                .AddEntry("Hardware", 2, "Section_Hardware")
                .AddEntry("Consulting", 2, "Section_Consulting")
                .AddEntry("Support Contracts", 2, "Section_SupportContracts")
                .AddEntry("Images & Charts", 1, "Section_Images")
                .AddEntry("Top Performers", 1, "Section_TopPerformers")
                .AddEntry("Detailed Daily Metrics", 1, "Section_DailyMetrics")
                .AddEntry("Summary & Recommendations", 1, "Section_Summary");

            // Get computed serial numbers for use in headings
            var tocEntries = toc.GetEntries();

            if (chkNavigation.IsChecked == true)
            {
                doc.AddTableOfContents(toc);
            }

            doc.AddPageBreak(); // Always page break after TOC

            // ===== Section 1: Executive Summary =====
            doc.AddHeading(tocEntries[0].SerialNumber + "  " + "Executive Summary", 1, bookmarkName: "Section_ExecutiveSummary");
            doc.AddParagraph(
                "This report provides a comprehensive overview of sales performance across all regions " +
                "for the current quarter. The data includes revenue figures, unit sales, growth metrics, " +
                "and year-over-year comparisons for each product category.");
            doc.AddParagraph("");
            if (autoSpacing) doc.AddSectionSeparator(); else doc.AddPageBreak();

            // ===== Section 2: Regional Sales Data =====
            doc.AddHeading(tocEntries[1].SerialNumber + "  " + "Regional Sales Data", 1, bookmarkName: "Section_RegionalSales");
            doc.AddParagraph("The following table summarizes sales data across all regions:");
            doc.AddParagraph("");

            var salesTable = new DocxTable()
                .SetBorders("single", 4, "4472C4")
                .AddHeaderRow("Region", "Q1 Revenue", "Q2 Revenue", "Q3 Revenue", "Q4 Revenue", "Total", "Growth %");

            string[] regions = { "North America", "Europe", "Asia Pacific", "Latin America", "Middle East", "Africa" };
            var random = new Random(42);

            foreach (var region in regions)
            {
                double q1 = random.Next(500000, 2000000);
                double q2 = random.Next(500000, 2000000);
                double q3 = random.Next(500000, 2000000);
                double q4 = random.Next(500000, 2000000);
                double total = q1 + q2 + q3 + q4;
                double growth = (random.NextDouble() * 30) - 5;

                salesTable.AddRow(region, $"${q1:N0}", $"${q2:N0}", $"${q3:N0}", $"${q4:N0}", $"${total:N0}", $"{growth:F1}%");
            }

            doc.AddTable(salesTable);
            if (autoSpacing) doc.AddSectionSeparator(); else doc.AddPageBreak();

            // ===== Section 3: Product Performance =====
            doc.AddHeading(tocEntries[2].SerialNumber + "  " + "Product Performance", 1, bookmarkName: "Section_ProductPerformance");

            string[] products = { "Software Licenses", "Cloud Services", "Hardware", "Consulting", "Support Contracts" };
            string[] productBookmarks = { "Section_SoftwareLicenses", "Section_CloudServices", "Section_Hardware", "Section_Consulting", "Section_SupportContracts" };

            for (int p = 0; p < products.Length; p++)
            {
                doc.AddHeading(tocEntries[3 + p].SerialNumber + "  " + products[p], 2, bookmarkName: productBookmarks[p]);
                doc.AddParagraph(
                    $"The {products[p]} division has shown significant progress this quarter. " +
                    $"Revenue targets were met with a {random.Next(85, 120)}% achievement rate. " +
                    $"Customer satisfaction remains high at {random.Next(80, 99)}%.");

                var productTable = new DocxTable()
                    .SetBorders("single", 2, "000000")
                    .AddHeaderRow("Month", "Units Sold", "Revenue", "Avg Price");

                string[] months = { "January", "February", "March", "April", "May", "June",
                                    "July", "August", "September", "October", "November", "December" };

                foreach (var month in months)
                {
                    int units = random.Next(100, 5000);
                    double revenue = units * random.Next(50, 500);
                    double avgPrice = revenue / units;
                    productTable.AddRow(month, units.ToString("N0"), $"${revenue:N0}", $"${avgPrice:F2}");
                }

                doc.AddTable(productTable);
                doc.AddParagraph("");
            }

            if (autoSpacing) doc.AddSectionSeparator(); else doc.AddPageBreak();

            // ===== Section 4: Images & Charts =====
            doc.AddHeading(tocEntries[8].SerialNumber + "  " + "Images & Charts", 1, bookmarkName: "Section_Images");

            if (_imagePaths.Count > 0)
            {
                doc.AddParagraph($"The following {_imagePaths.Count} image(s) have been included in this report:");
                doc.AddParagraph("");

                foreach (var imagePath in _imagePaths)
                {
                    doc.AddParagraph(Path.GetFileName(imagePath), bold: true, fontSize: 11);
                    doc.AddImage(imagePath, widthEmu: 4500000, heightEmu: 3000000);
                    doc.AddParagraph("");
                }
            }
            else
            {
                doc.AddParagraph("No images were provided for this report.", italic: true);
            }

            if (autoSpacing) doc.AddSectionSeparator(); else doc.AddPageBreak();

            // ===== Section 5: Top Performers =====
            doc.AddHeading(tocEntries[9].SerialNumber + "  " + "Top Performers", 1, bookmarkName: "Section_TopPerformers");
            doc.AddParagraph("The following employees have exceeded their quarterly targets:");

            var empTable = new DocxTable()
                .SetBorders("single", 4, "2E7D32")
                .AddHeaderRow("Rank", "Employee", "Department", "Target", "Achieved", "% of Target");

            string[] names = { "Alice Johnson", "Bob Smith", "Carol Williams", "David Brown",
                               "Eva Martinez", "Frank Lee", "Grace Chen", "Henry Wilson",
                               "Iris Patel", "Jack Thompson", "Karen Davis", "Leo Garcia",
                               "Maria Rodriguez", "Nathan Kim", "Olivia Taylor", "Peter Anderson" };

            for (int i = 0; i < names.Length; i++)
            {
                double target = random.Next(100000, 500000);
                double achieved = target * (1.0 + random.NextDouble() * 0.5);
                string[] depts = { "Sales", "Marketing", "Engineering", "Support" };

                empTable.AddRow(
                    (i + 1).ToString(), names[i], depts[random.Next(depts.Length)],
                    $"${target:N0}", $"${achieved:N0}", $"{(achieved / target * 100):F1}%");
            }

            doc.AddTable(empTable);
            if (autoSpacing) doc.AddSectionSeparator(); else doc.AddPageBreak();

            // ===== Section 6: Daily Metrics =====
            doc.AddHeading(tocEntries[10].SerialNumber + "  " + "Detailed Daily Metrics", 1, bookmarkName: "Section_DailyMetrics");
            doc.AddParagraph("Daily performance metrics for the past 90 days:");

            var metricsTable = new DocxTable()
                .SetBorders("single", 2, "666666")
                .AddHeaderRow("Date", "Visitors", "Conversions", "Revenue", "Avg Order");

            var startDate = DateTime.Now.AddDays(-90);
            for (int i = 0; i < 90; i++)
            {
                var date = startDate.AddDays(i);
                int visitors = random.Next(1000, 50000);
                int conversions = (int)(visitors * (random.NextDouble() * 0.05 + 0.01));
                double revenue = conversions * random.Next(20, 200);
                double avgOrder = conversions > 0 ? revenue / conversions : 0;

                metricsTable.AddRow(
                    date.ToString("yyyy-MM-dd"), visitors.ToString("N0"),
                    conversions.ToString("N0"), $"${revenue:N0}", $"${avgOrder:F2}");
            }

            doc.AddTable(metricsTable);
            if (autoSpacing) doc.AddSectionSeparator(); else doc.AddPageBreak();

            // ===== Section 7: Summary =====
            doc.AddHeading(tocEntries[11].SerialNumber + "  " + "Summary & Recommendations", 1, bookmarkName: "Section_Summary");
            doc.AddParagraph(
                "Based on the comprehensive data analysis presented in this report, the following " +
                "key recommendations are made for the upcoming quarter:", bold: true);
            doc.AddParagraph("");
            doc.AddParagraph("1. Increase investment in Asia Pacific region due to high growth potential.");
            doc.AddParagraph("2. Expand Cloud Services team to meet rising demand.");
            doc.AddParagraph("3. Implement new customer retention program for Support Contracts.");
            doc.AddParagraph("4. Review pricing strategy for Hardware division.");
            doc.AddParagraph("5. Recognize and reward top performers to maintain motivation.");
            doc.AddParagraph("");
            doc.AddHorizontalRule();
            doc.AddParagraph("End of Report", alignment: "center", italic: true);

            doc.Save(filePath);
        }

        private void GenerateLargePdfReport(string filePath)
        {
            var doc = new PdfDocument();
            doc.SetPageMargins(topPt: 72, bottomPt: 72, leftPt: 72, rightPt: 72);

            if (chkHeader.IsChecked == true)
                doc.SetHeader("Company Quarterly Report");

            if (chkFooter.IsChecked == true || chkPageNumbers.IsChecked == true)
                doc.SetFooter(chkFooter.IsChecked == true ? "Confidential" : null,
                              includePageNumbers: chkPageNumbers.IsChecked == true);

            // ===== Build TOC =====
            var toc = new PdfTableOfContents()
                .SetTitle("Table of Contents")
                .ShowSerialNumbers(true)
                .UseLinkStyle(true)
                .AddEntry("Executive Summary", 1, "Section_ExecutiveSummary")
                .AddEntry("Regional Sales Data", 1, "Section_RegionalSales")
                .AddEntry("Product Performance", 1, "Section_ProductPerformance")
                .AddEntry("Software Licenses", 2, "Section_SoftwareLicenses")
                .AddEntry("Cloud Services", 2, "Section_CloudServices")
                .AddEntry("Hardware", 2, "Section_Hardware")
                .AddEntry("Consulting", 2, "Section_Consulting")
                .AddEntry("Support Contracts", 2, "Section_SupportContracts")
                .AddEntry("Images & Charts", 1, "Section_Images")
                .AddEntry("Top Performers", 1, "Section_TopPerformers")
                .AddEntry("Detailed Daily Metrics", 1, "Section_DailyMetrics")
                .AddEntry("Summary & Recommendations", 1, "Section_Summary");

            var tocEntries = toc.GetEntries();

            // ===== Title Page =====
            doc.AddHeading("Quarterly Sales Report", 1);
            doc.AddParagraph("Generated on: " + DateTime.Now.ToString("MMMM dd, yyyy"), fontSize: 12);
            doc.AddParagraph("Department: Sales & Marketing", fontSize: 12);
            doc.AddHorizontalRule();
            doc.AddParagraph("");

            if (chkNavigation.IsChecked == true)
                doc.AddTableOfContents(toc);

            doc.AddPageBreak();

            // ===== Section 1: Executive Summary =====
            doc.AddHeading(tocEntries[0].SerialNumber + "  Executive Summary", 1, bookmarkName: "Section_ExecutiveSummary");
            doc.AddParagraph(
                "This report provides a comprehensive overview of sales performance across all regions " +
                "for the current quarter. The data includes revenue figures, unit sales, growth metrics, " +
                "and year-over-year comparisons for each product category.");
            doc.AddPageBreak();

            // ===== Section 2: Regional Sales Data =====
            doc.AddHeading(tocEntries[1].SerialNumber + "  Regional Sales Data", 1, bookmarkName: "Section_RegionalSales");
            doc.AddParagraph("The following table summarizes sales data across all regions:");

            var salesTable = new PdfTable()
                .SetBorders("single", 4, "4472C4")
                .AddHeaderRow("Region", "Q1 Revenue", "Q2 Revenue", "Q3 Revenue", "Q4 Revenue", "Total", "Growth %");

            string[] regions = { "North America", "Europe", "Asia Pacific", "Latin America", "Middle East", "Africa" };
            var random = new Random(42);
            foreach (var region in regions)
            {
                double q1 = random.Next(500000, 2000000); double q2 = random.Next(500000, 2000000);
                double q3 = random.Next(500000, 2000000); double q4 = random.Next(500000, 2000000);
                double total = q1 + q2 + q3 + q4; double growth = (random.NextDouble() * 30) - 5;
                salesTable.AddRow(region, $"${q1:N0}", $"${q2:N0}", $"${q3:N0}", $"${q4:N0}", $"${total:N0}", $"{growth:F1}%");
            }
            doc.AddTable(salesTable);
            doc.AddPageBreak();

            // ===== Section 3: Product Performance =====
            doc.AddHeading(tocEntries[2].SerialNumber + "  Product Performance", 1, bookmarkName: "Section_ProductPerformance");

            string[] products = { "Software Licenses", "Cloud Services", "Hardware", "Consulting", "Support Contracts" };
            string[] productBookmarks = { "Section_SoftwareLicenses", "Section_CloudServices", "Section_Hardware", "Section_Consulting", "Section_SupportContracts" };

            for (int p = 0; p < products.Length; p++)
            {
                doc.AddHeading(tocEntries[3 + p].SerialNumber + "  " + products[p], 2, bookmarkName: productBookmarks[p]);
                doc.AddParagraph($"The {products[p]} division has shown significant progress this quarter. Revenue targets were met.");

                var productTable = new PdfTable()
                    .SetBorders("single", 2, "000000")
                    .AddHeaderRow("Month", "Units Sold", "Revenue", "Avg Price");
                string[] months = { "January", "February", "March", "April", "May", "June",
                                    "July", "August", "September", "October", "November", "December" };
                foreach (var month in months)
                {
                    int units = random.Next(100, 5000);
                    double rev = units * random.Next(50, 500);
                    productTable.AddRow(month, units.ToString("N0"), $"${rev:N0}", $"${rev / units:F2}");
                }
                doc.AddTable(productTable);
            }
            doc.AddPageBreak();

            // ===== Section 4: Images & Charts =====
            doc.AddHeading(tocEntries[8].SerialNumber + "  Images & Charts", 1, bookmarkName: "Section_Images");
            if (_imagePaths.Count > 0)
            {
                doc.AddParagraph($"{_imagePaths.Count} image(s) included in this report:");
                foreach (var imgPath in _imagePaths)
                {
                    doc.AddParagraph(Path.GetFileName(imgPath), bold: true);
                    byte[] jpegBytes = LoadImageAsJpeg(imgPath);
                    if (jpegBytes != null)
                        doc.AddImage(jpegBytes, widthPt: 400f, alignment: "center");
                    doc.AddLineBreak();
                }
            }
            else
            {
                doc.AddParagraph("No images were provided for this report.");
            }
            doc.AddPageBreak();

            // ===== Section 5: Top Performers =====
            doc.AddHeading(tocEntries[9].SerialNumber + "  Top Performers", 1, bookmarkName: "Section_TopPerformers");
            doc.AddParagraph("The following employees have exceeded their quarterly targets:");

            var empTable = new PdfTable()
                .SetBorders("single", 4, "2E7D32")
                .AddHeaderRow("Rank", "Employee", "Department", "Target", "Achieved", "% of Target");
            string[] names = { "Alice Johnson", "Bob Smith", "Carol Williams", "David Brown",
                               "Eva Martinez", "Frank Lee", "Grace Chen", "Henry Wilson",
                               "Iris Patel", "Jack Thompson", "Karen Davis", "Leo Garcia",
                               "Maria Rodriguez", "Nathan Kim", "Olivia Taylor", "Peter Anderson" };
            string[] depts = { "Sales", "Marketing", "Engineering", "Support" };
            for (int i = 0; i < names.Length; i++)
            {
                double target = random.Next(100000, 500000);
                double achieved = target * (1.0 + random.NextDouble() * 0.5);
                empTable.AddRow((i + 1).ToString(), names[i], depts[random.Next(depts.Length)],
                    $"${target:N0}", $"${achieved:N0}", $"{achieved / target * 100:F1}%");
            }
            doc.AddTable(empTable);
            doc.AddPageBreak();

            // ===== Section 6: Daily Metrics =====
            doc.AddHeading(tocEntries[10].SerialNumber + "  Detailed Daily Metrics", 1, bookmarkName: "Section_DailyMetrics");
            doc.AddParagraph("Daily performance metrics for the past 90 days:");

            var metricsTable = new PdfTable()
                .SetBorders("single", 2, "666666")
                .AddHeaderRow("Date", "Visitors", "Conversions", "Revenue", "Avg Order");
            var startDate = DateTime.Now.AddDays(-90);
            for (int i = 0; i < 90; i++)
            {
                var date = startDate.AddDays(i);
                int visitors = random.Next(1000, 50000);
                int conversions = (int)(visitors * (random.NextDouble() * 0.05 + 0.01));
                double revenue = conversions * random.Next(20, 200);
                metricsTable.AddRow(date.ToString("yyyy-MM-dd"), visitors.ToString("N0"),
                    conversions.ToString("N0"), $"${revenue:N0}", $"${(conversions > 0 ? revenue / conversions : 0):F2}");
            }
            doc.AddTable(metricsTable);
            doc.AddPageBreak();

            // ===== Section 7: Summary =====
            doc.AddHeading(tocEntries[11].SerialNumber + "  Summary & Recommendations", 1, bookmarkName: "Section_Summary");
            doc.AddParagraph("Based on the comprehensive data analysis presented in this report, the following key recommendations are made for the upcoming quarter:", bold: true);
            doc.AddParagraph("1. Increase investment in Asia Pacific region due to high growth potential.");
            doc.AddParagraph("2. Expand Cloud Services team to meet rising demand.");
            doc.AddParagraph("3. Implement new customer retention program for Support Contracts.");
            doc.AddParagraph("4. Review pricing strategy for Hardware division.");
            doc.AddParagraph("5. Recognize and reward top performers to maintain motivation.");
            doc.AddHorizontalRule();
            doc.AddParagraph("End of Report", alignment: "center");

            doc.Save(filePath);
        }
            /// <summary>
            /// Loads any WPF-supported image and returns it as a JPEG byte array
            /// suitable for embedding in a PDF /DCTDecode XObject.
            /// </summary>
            private static byte[] LoadImageAsJpeg(string path)
            {
                try
                {
                    // Already a JPEG — return raw bytes directly
                    string ext = Path.GetExtension(path).ToLowerInvariant();
                    if (ext == ".jpg" || ext == ".jpeg")
                        return File.ReadAllBytes(path);

                    // Other formats: decode then re-encode to JPEG via WPF
                    var bmp = new BitmapImage();
                    bmp.BeginInit();
                    bmp.UriSource = new Uri(path, UriKind.Absolute);
                    bmp.CacheOption = BitmapCacheOption.OnLoad;
                    bmp.EndInit();
                    bmp.Freeze();

                    using var ms = new MemoryStream();
                    var encoder = new JpegBitmapEncoder { QualityLevel = 85 };
                    encoder.Frames.Add(BitmapFrame.Create(bmp));
                    encoder.Save(ms);
                    return ms.ToArray();
                }
                catch
                {
                    return Array.Empty<byte>();
                }
            }
        }
    }

