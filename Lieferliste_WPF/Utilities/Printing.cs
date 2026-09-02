using El2Core.Models;
using El2Core.Utils;
using El2Core.ViewModelBase;
using Lieferliste_WPF.Views;
using ModulePlanning.Planning;
using System;
using System.Collections.Generic;
using System.Drawing.Printing;
using System.Globalization;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Printing;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Resources;
using System.Windows.Xps.Packaging;
using System.Windows.Xps.Serialization;
using Thickness = System.Windows.Thickness;
using ValidationResult = System.Printing.ValidationResult;

namespace Lieferliste_WPF.Utilities
{
    [System.Runtime.Versioning.SupportedOSPlatform("windows10.0")]
    public interface IPrinting
    {
        void PrintPreview(PlanMachine planMachine, PrintTicket ticket);
        void Print(PlanMachine planMachine);
        void PrintPreview(List<Vorgang> vorgangs, PrintTicket ticket);
        void Print(List<Vorgang> vorgangs);
        void PrintPreview(IViewModel projects, PrintTicket ticket);

    }
    public class PrintingProxy : IPrinting
    {
        public void Print(PlanMachine planMachine)
        {
            throw new NotImplementedException();
        }

        public void Print(List<Vorgang> vorgangs)
        {
            throw new NotImplementedException();
        }
        /// <summary>
        /// Print Preview for a PlanMachine with a given PrintTicket
        /// </summary>
        /// <param name="planMachine"></param>
        /// <param name="ticket"></param>
        public void PrintPreview(PlanMachine planMachine, PrintTicket ticket)
        {
            Printing.DoPrintPreview
                (planMachine.Name,
                planMachine.InventNo,
                planMachine.Description,
                planMachine.Processes.ToList(),
                ticket);
        }

        public void PrintPreview(IViewModel viewModel, PrintTicket ticket)
        {

        }

        public void PrintPreview(List<Vorgang> vorgangs, PrintTicket ticket)
        {
            throw new NotImplementedException();
        }
    }
    internal static class Printing
    {
        /// <summary>
        /// Prints the given FlowDocument using the specified PrintTicket and description.
        /// </summary>
        /// <param name="document"></param>
        /// <param name="ticket"></param>
        /// <param name="description"></param>
        public static void DoThePrint(FlowDocument document, PrintTicket ticket, string description)
        {
            // Clone the source document's content into a new FlowDocument.
            // This is because the pagination for the printer needs to be
            // done differently than the pagination for the displayed page.
            // We print the copy, rather that the original FlowDocument.
            System.IO.MemoryStream s = new System.IO.MemoryStream();
            TextRange source = new TextRange(document.ContentStart, document.ContentEnd);
            source.Save(s, DataFormats.Xaml);
            FlowDocument copy = new FlowDocument();
            TextRange dest = new TextRange(copy.ContentStart, copy.ContentEnd);
            dest.Load(s, DataFormats.Xaml);


            PrintQueue queue = new PrintServer().GetPrintQueue(new PrinterSettings().PrinterName);
            queue.CurrentJobSettings.Description = description;
            queue.CurrentJobSettings.CurrentPrintTicket = ticket;

            ValidationResult result = queue.MergeAndValidatePrintTicket(queue.UserPrintTicket, ticket);
            System.Windows.Xps.XpsDocumentWriter docWriter = System.Printing.PrintQueue.CreateXpsDocumentWriter(queue);

            DocumentPaginator paginator = ((IDocumentPaginatorSource)document).DocumentPaginator;

            document.ColumnWidth = double.PositiveInfinity;
            docWriter.Write(paginator, ticket);

        }
        private static FlowDocument CreateFlowDocument(string Name, string? Second, string? Description, List<Vorgang> proces, PrintTicket ticket)
        {


            var printDlg = new PrintDialog();
            printDlg.PrintTicket = ticket;
            StreamResourceInfo BitmapStreamSourceInfo;

            FlowDocument fd = new FlowDocument();
            Paragraph logo = new Paragraph(new Run("COS-Lieferliste"));
            logo.FontSize = 20;
            logo.FontWeight = FontWeights.Bold;
            logo.Padding = new Thickness(0.0, 15.0, 0.0, 0.0);
            BitmapImage bitmap = new BitmapImage(); 
            BitmapStreamSourceInfo = Application.GetResourceStream(new Uri("..\\Images\\BOSCH2.jpg", UriKind.Relative));

            bitmap.BeginInit();
            bitmap.StreamSource = BitmapStreamSourceInfo.Stream;
            bitmap.EndInit();
            Image l = new Image() { Source = bitmap };
            Figure figureLogo = new Figure();
            figureLogo.Width = new FigureLength(200);
            figureLogo.Height = new FigureLength(120);
            figureLogo.Background = Brushes.GhostWhite;
            figureLogo.HorizontalAnchor = FigureHorizontalAnchor.PageLeft;
            figureLogo.Blocks.Add(new BlockUIContainer(l));

            figureLogo.Blocks.Add(logo);

            Paragraph p1 = new Paragraph();
            p1.Inlines.Add(figureLogo);
            p1.Inlines.Add(new Run(DateTime.Now.ToString("ddd, dd/MM/yyyy HH:mm")));

            fd.PageWidth = printDlg.PrintableAreaWidth;
            fd.PageHeight = printDlg.PrintableAreaHeight;
            fd.PagePadding = new Thickness(96 / 3);
            double printSizeWidth = fd.PageWidth - 96 / 3 * 2;
            p1.FontStyle = FontStyles.Normal;
            p1.FontFamily = new FontFamily("Microsoft Sans Serif");
            p1.FontSize = 12;
            fd.Blocks.Add(p1);
            StringBuilder stringBuilder = new StringBuilder();
            stringBuilder.Append(Name).Append(' ').AppendLine(Second).Append(Description);
            Paragraph p = new Paragraph(new Run(stringBuilder.ToString()));
            p.FontStyle = FontStyles.Normal;
            p.FontWeight = FontWeights.Bold;
            p.FontFamily = new FontFamily("Segoe UI");
            p.FontSize = 18;
            fd.Blocks.Add(p);

            Table table = new Table();
            TableRowGroup tableRowGroup = new TableRowGroup();
            TableRow r = new TableRow();

            fd.BringIntoView();
            fd.ColumnWidth = 800;
            fd.TextAlignment = TextAlignment.Center;
            table.CellSpacing = 1;


            var headerList = proces.OrderBy(x => x.SortPos).Select(x => new
            {
                x.Aid,
                ProcessingUom = string.Format("{0:d4}", x.Vnr),
                x.AidNavigation.Material,
                x.AidNavigation.MaterialNavigation?.Bezeichng,
                QuantString = x.AidNavigation.Quantity.ToString(),
                x.Text,
                StartKW = string.Format("{0} KW{1}", x.SpaetStart.GetValueOrDefault().ToString("dd/MM/yy"),
                        CultureInfo.CurrentCulture.Calendar.GetWeekOfYear(x.SpaetStart.GetValueOrDefault(), CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday)),
                SpaetEnd = x.SpaetEnd.GetValueOrDefault().ToShortDateString(),
                BemTL = (x.BemT != null && x.BemT.Split((char)29).Length == 2) ? x.BemT.Split((char)29)[1] : string.Empty,
                TimeLength = string.Format("{0:F2}h", (((x.Beaze == null) ? 0 : x.Beaze) + ((x.Rstze == null) ? 0 : x.Rstze)) / 60)
            }).ToArray();

            int i = 0;
            foreach (var pr in headerList.First().GetType().GetProperties())
            {
                TableColumn tabCol = new TableColumn();
                string head;
                switch (pr.Name)
                {
                    case "Aid": head = "Auftrags-\nnummer"; tabCol.Width = new GridLength(printSizeWidth * 0.08); break;
                    case "ProcessingUom": head = "Vorgang"; tabCol.Width = new GridLength(printSizeWidth * 0.035); break;
                    case "Material": head = "Material"; tabCol.Width = new GridLength(printSizeWidth * 0.09); break;
                    case "Bezeichng": head = "Bezeichnung"; tabCol.Width = new GridLength(printSizeWidth * 0.13); break;
                    case "QuantString": head = "Stk."; tabCol.Width = new GridLength(printSizeWidth * 0.035); break;
                    case "Text": head = "Kurztext"; tabCol.Width = new GridLength(printSizeWidth * 0.19); break;
                    case "StartKW": head = "SpätStart"; tabCol.Width = new GridLength(printSizeWidth * 0.06); break;
                    case "SpaetEnd": head = "SpätEnd"; tabCol.Width = new GridLength(printSizeWidth * 0.07); break;
                    case "BemTL": head = "Bemerkung Teamleiter"; tabCol.Width = new GridLength(printSizeWidth * 0.24); break;
                    case "TimeLength": head = "Dauer"; tabCol.Width = new GridLength(printSizeWidth * 0.06); break;
                    default: head = "not Valid"; break;
                }
                r.Cells.Add(new TableCell(new Paragraph(new Run(head)))
                {
                    BorderThickness = new Thickness(1, 0, 0, 1),
                    BorderBrush = Brushes.Black,
                    Background = Brushes.DarkGray,
                    Foreground = Brushes.White
                });
                r.Cells[i].Padding = new Thickness(2);
                table.Columns.Add(tabCol);

                ++i;
            }
            tableRowGroup.Rows.Add(r);
            table.RowGroups.Add(tableRowGroup);
            foreach (var row in headerList)
            {

                table.BorderBrush = Brushes.Gray;
                table.BorderThickness = new Thickness(1, 1, 0, 0);
                table.FontStyle = FontStyles.Normal;
                table.FontFamily = new FontFamily("Arial");
                table.FontSize = 13;
                tableRowGroup = new TableRowGroup();
                r = new TableRow();
                i = 0;
                foreach (var property in row.GetType().GetProperties())
                {

                    r.Cells.Add(new TableCell(new Paragraph(new Run(property.GetValue(row)?.ToString()))));
                    r.Cells[i].Padding = new Thickness(1);
                    r.Cells[i].BorderBrush = Brushes.DarkGray;
                    r.Cells[i].BorderThickness = new Thickness(0, 0, 1, 1);
                    if (i == 8) r.Cells[i].TextAlignment = TextAlignment.Left;
                    ++i;
                }

                tableRowGroup.Rows.Add(r);
                table.RowGroups.Add(tableRowGroup);
            }
            fd.Blocks.Add(table);

            return fd;

        }

        internal static void DoPrintPreview(string name, string? second, string? description, List<Vorgang> proces, PrintTicket ticket)
        {

            var doc = CreateFlowDocument(name, second, description, proces, ticket);
            var fix = Get_Fixed_From_FlowDoc(doc, ticket);

            var windows = new PrintWindow(fix);
            windows.ShowDialog();

        }
        private static string _previewWindowXaml =
    @"<Window
        xmlns ='http://schemas.microsoft.com/netfx/2007/xaml/presentation'
        xmlns:x ='http://schemas.microsoft.com/winfx/2006/xaml'
        Title ='Print Preview - @@TITLE'
        Height ='200' Width ='300'
        WindowStartupLocation ='CenterOwner'>
                      <DocumentViewer Name='dv1'/>
     </Window>";
        /// <summary>
        /// Previews a window with the given useful data.
        /// The window is created using a Grid and a FixedDocument, and the document is displayed in a DocumentViewer.
        /// </summary>
        /// <param name="usefulData"></param>
        public static void PreviewWindowXaml(object usefulData)
        {
            Grid grid;
            grid = new Grid
            {
                //ItemsSource = usefulData
            };

            FixedDocument fixedDoc = new FixedDocument();
            PageContent pageContent = new PageContent();
            FixedPage fixedPage = new FixedPage();

            //Create first page of document
            fixedPage.Children.Add(grid);
            ((System.Windows.Markup.IAddChild)pageContent).AddChild(fixedPage);
            fixedDoc.Pages.Add(pageContent);
            //Create any other required pages here
            DocumentViewer documentViewer1 = new DocumentViewer();
            //View the document
            documentViewer1.Document = fixedDoc;
        }
        //public static void DoPreview(string title)
        //{
        //    string fileName = System.IO.Path.GetRandomFileName();


        //    FlowDocumentScrollViewer visual = (FlowDocumentScrollViewer)_parent.FindName("fdsv1");

        //    try
        //    {
        //        // write the XPS document
        //        using (XpsDocument doc = new XpsDocument(fileName, FileAccess.ReadWrite))
        //        {
        //            XpsDocumentWriter writer = XpsDocument.CreateXpsDocumentWriter(doc);
        //            writer.Write(visual);
        //        }

        //        // Read the XPS document into a dynamically generated
        //        // preview Window 
        //        using (XpsDocument doc = new XpsDocument(fileName, FileAccess.Read))
        //        {
        //            FixedDocumentSequence fds = doc.GetFixedDocumentSequence();

        //            string s = _previewWindowXaml;
        //            s = s.Replace("@@TITLE", title.Replace("'", "&apos;"));

        //            using (var reader = new System.Xml.XmlTextReader(new StringReader(s)))
        //            {
        //                Window preview = System.Windows.Markup.XamlReader.Load(reader) as Window;

        //                DocumentViewer dv1 = LogicalTreeHelper.FindLogicalNode(preview, "dv1") as DocumentViewer;
        //                dv1.Document = fds as IDocumentPaginatorSource;


        //                preview.ShowDialog();
        //            }
        //        }
        //    }
        //    finally
        //    {
        //        if (File.Exists(fileName))
        //        {
        //            try
        //            {
        //                File.Delete(fileName);
        //            }
        //            catch
        //            {
        //            }
        //        }
        //    }
        //}

        /// <summary>
        /// Converts a FlowDocument to a FixedDocument using the specified PrintTicket.
        /// </summary>
        /// <param name="flowDoc"></param>
        /// <param name="ticket"></param>
        /// <returns></returns>
        public static FixedDocument Get_Fixed_From_FlowDoc(FlowDocument flowDoc, PrintTicket ticket)
        {
            var fixedDocument = new FixedDocument();
            try
            {
                var pdlgPrint = new PrintDialog();
                pdlgPrint.PrintTicket = ticket;
                DocumentPaginator dpPages = (DocumentPaginator)((IDocumentPaginatorSource)flowDoc).DocumentPaginator;
                dpPages.ComputePageCount();
                PrintCapabilities capabilities = pdlgPrint.PrintQueue.GetPrintCapabilities(pdlgPrint.PrintTicket);

                for (int iPages = 0; iPages < dpPages.PageCount; iPages++)
                {
                    var page = dpPages.GetPage(iPages);
                    var pageContent = new PageContent();
                    var fixedPage = new FixedPage();

                    Canvas canvas = new Canvas();

                    VisualBrush vb = new VisualBrush(page.Visual);
                    vb.Stretch = Stretch.None;
                    vb.AlignmentX = AlignmentX.Left;
                    vb.AlignmentY = AlignmentY.Top;
                    vb.ViewboxUnits = BrushMappingMode.Absolute;
                    vb.TileMode = TileMode.None;
                    vb.Viewbox = new Rect(0, 0, capabilities.PageImageableArea.ExtentWidth, capabilities.PageImageableArea.ExtentHeight);

                    FixedPage.SetLeft(canvas, 0);
                    FixedPage.SetTop(canvas, 0);

                    canvas.Width = capabilities.PageImageableArea.ExtentWidth;
                    canvas.Height = capabilities.PageImageableArea.ExtentHeight;
                    canvas.Background = vb;

                    fixedPage.Children.Add(canvas);
                    fixedPage.Width = pdlgPrint.PrintableAreaWidth;
                    fixedPage.Height = pdlgPrint.PrintableAreaHeight;
                    fixedPage.HorizontalAlignment = HorizontalAlignment.Stretch;
                    pageContent.Child = fixedPage;

                    fixedDocument.Pages.Add(pageContent);
                }

            }
            catch (Exception)
            {
                throw;
            }
            return fixedDocument;
        }
        /// <summary>
        /// Creates a FlowDocument for a given Vorgang,
        /// including details such as Auftrag, Vorgang number, Material, Bemerkung, and Start messen time.
        /// </summary>
        /// <param name="vorgang"></param>
        /// <returns></returns>
        public static FlowDocument CreateKlimaDocument(Vorgang vorgang)
        {
            FlowDocument document = new FlowDocument();

            document.PageWidth = 768;
            document.PageHeight = 554;
            document.PagePadding = new Thickness(96 / 3);

            Paragraph p1 = new Paragraph();
            p1.FontSize = 16;
            p1.FontFamily = SystemFonts.CaptionFontFamily;

            p1.Inlines.Add(new Run("Auftrag: "));
            p1.Inlines.Add(new Bold(new Underline(new Run(vorgang.Aid))));

            Paragraph p2 = new Paragraph();
            p2.FontSize = 16;
            p2.FontFamily = SystemFonts.CaptionFontFamily;

            p2.Inlines.Add(new Run(string.Format("Vorgang: {0:D4} -- {1}", vorgang.Vnr, vorgang.Text)));

            Paragraph p3 = new Paragraph();
            p3.FontSize = 16;
            p3.FontFamily = SystemFonts.CaptionFontFamily;

            p3.Inlines.Add(new Run("Material: "));

            p3.Inlines.Add(new Bold(new Run(string.Format("{0} {1}", vorgang.AidNavigation.Material,
                vorgang.AidNavigation.MaterialNavigation?.Bezeichng))));
            Figure figure = new Figure();
            figure.HorizontalAnchor = FigureHorizontalAnchor.ContentRight;
            Paragraph p4 = new Paragraph();
            Run r = new Run(string.Format("Spät. Enddatum: {0}", vorgang.SpaetEnd?.ToString("dd/MM/yyyy")));
            p4.Inlines.Add(r);
            figure.Blocks.Add(p4);
            p1.Inlines.Add(figure);
            Paragraph p5 = new Paragraph();
            p5.FontSize = 16;
            p5.FontFamily = SystemFonts.CaptionFontFamily;
            p5.Inlines.Add(new Run("Bemerkung: "));
            p5.Inlines.Add(new Run(vorgang.BemT));

            var h = int.Parse(RuleInfo.Rules["ClimaticWaitTime"].RuleValue);
            var hh = (vorgang.KlimaPrint.HasValue) ? vorgang.KlimaPrint.Value.AddHours(h).ToString("dd/MM/yy HH:mm:ss") : null;

            Paragraph p6 = new Paragraph(new Run(string.Format("Start messen frühestens {0}", hh)));
            p6.FontSize = 28;
            p6.FontFamily = SystemFonts.MessageFontFamily;
            p6.FontWeight = FontWeights.ExtraBold;
            p6.TextAlignment = TextAlignment.Center;
            p6.TextDecorations = TextDecorations.Underline;
            p6.BorderBrush = Brushes.Black;
            p6.BorderThickness = new Thickness(2);

            Section section = new Section();
            section.Blocks.Add(p6);

            document.Blocks.Add(p1);
            document.Blocks.Add(p2);
            document.Blocks.Add(p3);
            document.Blocks.Add(p5);
            document.Blocks.Add(section);

            return document;
        }
        //public static void FlexPrint()
        //{
        //    // Dimensions are 1/96th of an inch, socalled Device Independent Pixels (matches 96 DPI, which is normal?)
        //    var linerMargin = 4.7244094488189;  // 1.25 mm
        //    var magicOffset = 20;               // 5.29 mm

        //    // The document

        //    var doc = new FixedDocument();
        //    doc.PrintTicket = new PrintTicket();
        //    var printTicket = (PrintTicket)doc.PrintTicket;
        //    printTicket.PageMediaSize = new PageMediaSize(PageMediaSizeName.Unknown, size.Width.Value, size.Height.Value);

        //    // New page

        //    var page = new FixedPage
        //    {
        //        Width = labelSize.Width + magicOffset + 2 * linerMargin,
        //        Height = labelSize.Height,
        //        Margin = new Thickness(magicOffset + linerMargin, 0, linerMargin, 0)
        //    };

        //    // Page content

        //    var border = new Border
        //    {
        //        Width = labelSize.Width,
        //        Height = labelSize.Height,
        //        BorderBrush = Brushes.Black,
        //        BorderThickness = new Thickness(4)
        //    };

        //    FixedPage.SetLeft(border, 0);
        //    FixedPage.SetTop(border, 0);
        //    page.Children.Add(border);

        //    // Next item on page

        //    var tb = new TextBlock
        //    {
        //        Text = "Hello World!",
        //        FontFamily = new FontFamily("Algerian Regular"),
        //        FontSize = 30,
        //        HorizontalAlignment = HorizontalAlignment.Left,
        //        VerticalAlignment = VerticalAlignment.Top
        //    };

        //    var content = new Border
        //    {
        //        Child = tb,
        //        BorderBrush = Brushes.Black,
        //        BorderThickness = new Thickness(1)
        //    };

        //    FixedPage.SetLeft(content, 10);
        //    FixedPage.SetTop(content, 10);
        //    page.Children.Add(content);

        //    // Next item on page

        //    var tb2 = new TextBlock
        //    {
        //        Text = "SomethingElse",
        //        FontFamily = new FontFamily("Segoe UI"),
        //        FontSize = 20,
        //        HorizontalAlignment = HorizontalAlignment.Left,
        //        VerticalAlignment = VerticalAlignment.Top
        //    };

        //    var content2 = new Border
        //    {
        //        Child = tb2,
        //        BorderBrush = Brushes.Black,
        //        BorderThickness = new Thickness(1)
        //    };

        //    FixedPage.SetLeft(content2, 20);
        //    FixedPage.SetTop(content2, 55);
        //    page.Children.Add(content2);

        //    //// Add the page to the document

        //    //var pc = new PageContent();     // Old way of doing it (.NET 3.5 and earlier)
        //    //((IAddChild)pc).AddChild(page); //
        //    //doc.Pages.Add(pc);              //

        //    var pc = new PageContent();
        //    pc.Child = page;
        //    doc.Pages.Add(pc);

        //    // Send it to the printer   

        //    FlexPrinter.Print(printerName, labelSize, doc);
        //}

        /// <summary>
        /// Saves the given FlowDocument as an XPS file at the specified path.
        /// </summary>
        /// <param name="path"></param>
        /// <param name="document"></param>
        public static void SaveAsXps(string path, FlowDocument document)
        {
            using (Package package = Package.Open(path, FileMode.Create))
            {
                using (var xpsDoc = new XpsDocument(
                    package, CompressionOption.Maximum))
                {
                    var xpsSm = new XpsSerializationManager(
                        new XpsPackagingPolicy(xpsDoc), false);
                    DocumentPaginator dp =
                        ((IDocumentPaginatorSource)document).DocumentPaginator;
                    xpsSm.SaveAsXaml(dp);
                }
            }
        }

    }
}
