using Aspose.Cells;
using Aspose.Pdf;
using Aspose.Pdf.Text;
using FileConversion.Entity;
using FileConversion.Service.ServiceInterface;
using FileConversion.Utils.Constants;

namespace FileConversion.Service.ServiceImplementation
{
    public class ReportGeneratorImplementation : ReportGeneratorInterface
    {
        public ReportGenerationResponseMessage CreateReport(List<IFormFile> files)
        {

            var invoices = new List<Invoice>();
            var previewReportPath = new List<string>();

            var generatedPdfPaths = new List<string>();

            try
            {

                if (files == null || files.Count == 0)
                {
                    return new ReportGenerationResponseMessage { Success = false, Message = "Files are not provided" };
                }

                foreach (var file in files)
                {
                    if (file.Length > 0)
                    {
                        var filePath = Path.GetTempFileName();
                        using (var stream = new FileStream(filePath, FileMode.Create))
                        {
                            file.CopyTo(stream);
                        }

                        if (file.FileName.EndsWith(".xlsx") || file.FileName.EndsWith(".xls"))
                        {
                            ExtractDataFromFiles(filePath, invoices, @"C:\Users\i44375\source\repos\FileConversion\Utils\FileConverter\GeneratedReports");
                        }
                        else
                        {
                            return new ReportGenerationResponseMessage { Success = false, Message = "Only Excel files are supported" };
                        }
                        if (invoices.Count > 0)
                        {
                            SaveInvoiceInPdf(invoices, Constants.invoiceReportStoragePath);
                        }

                        string previewPdfPath = Path.Combine(Constants.invoiceReportStoragePath, $"Invoice_" + invoices[0].InvoiceNumber.ToString() + "_Report.pdf");
                        previewReportPath.Add(previewPdfPath);
                        System.IO.File.Delete(filePath);
                    }

                }

                foreach (var invoice in invoices)
                {
                    string outputPdfPath = Path.Combine(Constants.invoiceReportStoragePath, $"Invoice_" + (invoice.InvoiceNumber).ToString() + "_Report.pdf");
                    generatedPdfPaths.Add(outputPdfPath);
                }
                return new ReportGenerationResponseMessage { Success = true, Message = Message.successfullySavePdfs, Data = generatedPdfPaths, PreviewData = previewReportPath };
            }
            catch (ApplicationException ex)
            {
                Console.WriteLine("Index error: ", ex.Message);
                return new ReportGenerationResponseMessage
                {
                    Success = false,
                    Message = Message.unmatchedInputFileForTemplate
                };
            }
            catch (Exception ex)
            {
                Console.WriteLine("An unexpected error occured:", ex.Message);
                return new ReportGenerationResponseMessage { Success = false, Message = Message.unexpectedErrorWhileProcessingFiles };
            }
        }

        private void ExtractDataFromFiles(string inputFilePath, List<Invoice> invoices, string outputFilePath)
        {
            Workbook workbook = new Workbook(inputFilePath);
            var existingInvoice = new Invoice();
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                for (int i = 1; i <= worksheet.Cells.MaxDataRow; i++)
                {
                    try
                    {



                        int invoiceNumber = worksheet.Cells[i, 0].IntValue;
                        DateTime invoiceDate = worksheet.Cells[i, 1].DateTimeValue;
                        string customerName = worksheet.Cells[i, 2].StringValue;
                        string billingAddress = worksheet.Cells[i, 3].StringValue;
                        string description = worksheet.Cells[i, 4].StringValue;
                        int quantity = worksheet.Cells[i, 5].IntValue;
                        int unitPrice = worksheet.Cells[i, 6].IntValue;
                        double totalPrice = worksheet.Cells[i, 7].DoubleValue;
                        double taxRate = worksheet.Cells[i, 8].DoubleValue;
                        float taxAmount = worksheet.Cells[i, 9].FloatValue;
                        float totalAmount = worksheet.Cells[i, 10].FloatValue;
                        double contactNumber = worksheet.Cells[i, 11].DoubleValue;
                        string email = worksheet.Cells[i, 12].StringValue;

                        var product = new Product
                        {
                            Description = description,
                            Quantity = quantity,
                            UnitPrice = unitPrice,
                            TotalPrice = totalPrice
                        };

                        existingInvoice = invoices.Find(inv => inv.InvoiceNumber == invoiceNumber);

                        if (existingInvoice != null)
                        {
                            existingInvoice.Products.Add(product);
                        }
                        else
                        {
                            var invoice = new Invoice
                            {
                                InvoiceNumber = invoiceNumber,
                                InvoiceDate = invoiceDate,
                                CustomerName = customerName,
                                BillingAddress = billingAddress,
                                Products = new List<Product> { product },
                                TotalPrice = totalPrice,
                                ContactNumber = contactNumber,
                                Email = email
                            };

                            invoices.Add(invoice);
                        }
                    }
                    catch (CellsException ex)
                    {
                        Console.WriteLine($"Error processing row {i}: {ex.Message}");
                        throw new ApplicationException($"Error processing row {i}: {ex.Message}", ex);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Unexpected error processing row {i}: {ex.Message}");
                        throw new ApplicationException($"Unexpected error processing row {i}: {ex.Message}", ex);
                    }
                }

            }
        }

        private void SaveInvoiceInPdf(List<Invoice> invoices, string outputFilePath)
        {
            foreach (var invoice in invoices)
            {
                Document document = new Document();

                Aspose.Pdf.Page page = document.Pages.Add();
                page.SetPageSize(PageSize.A4.Width, PageSize.A4.Height);

                page.PageInfo.Margin = new MarginInfo { Top = 30, Bottom = 20, Left = 50, Right = 50 };

                createHeader(page);

                createUserInfo(page, invoice);

                createInvoiceTable(page, invoice);

                double totalInvoiceAmount = invoice.Products.Sum(p => p.TotalPrice);


                createTotalPrice(page, totalInvoiceAmount);



                createFooter(page);


                string outputPdfPath = Path.Combine(outputFilePath, "Invoice_" + (invoice.InvoiceNumber).ToString() + "_Report.pdf");

                document.Save(outputPdfPath);
            }

        }

        private void createTotalPrice(Aspose.Pdf.Page page, double subTotalAmount)
        {
            double taxableAmount = subTotalAmount * 0.1;
            double totalAmount = subTotalAmount + taxableAmount;

            TextFragment subTotalChargeLabel = new TextFragment($"Sub Total     : {subTotalAmount.ToString("C")}");
            subTotalChargeLabel.TextState.FontSize = 12;
            subTotalChargeLabel.HorizontalAlignment = HorizontalAlignment.Right;
            subTotalChargeLabel.Margin.Top = 15;
            subTotalChargeLabel.Margin.Bottom = 5;
            page.Paragraphs.Add(subTotalChargeLabel);

            TextFragment taxCharge = new TextFragment($"Tax 10%     : {taxableAmount.ToString("C")}");
            taxCharge.TextState.FontSize = 12;
            taxCharge.HorizontalAlignment = HorizontalAlignment.Right;
            taxCharge.Margin.Top = 15;
            taxCharge.Margin.Bottom = 5;
            page.Paragraphs.Add(taxCharge);

            TextFragment totalChargeLabel = new TextFragment($"TOTAL     : {totalAmount.ToString("C")}");
            totalChargeLabel.TextState.FontStyle = FontStyles.Bold;
            totalChargeLabel.TextState.FontSize = 12;
            totalChargeLabel.HorizontalAlignment = HorizontalAlignment.Right;
            totalChargeLabel.Margin.Top = 15;
            totalChargeLabel.Margin.Bottom = 5;
            page.Paragraphs.Add(totalChargeLabel);
        }



        private void createInvoiceTable(Aspose.Pdf.Page page, Invoice invoice)
        {
            Aspose.Pdf.Table productsTable = new Aspose.Pdf.Table();
            productsTable.ColumnWidths = "150 100 100 100 100 100 100";

            productsTable.Border = new Aspose.Pdf.BorderInfo(Aspose.Pdf.BorderSide.Bottom, 1f, Aspose.Pdf.Color.LightBlue);

            page.Paragraphs.Add(productsTable);

            AddProductHeaders(productsTable);

            foreach (var product in invoice.Products)
            {
                AddProductDetails(productsTable, product.Description, product.Quantity.ToString(), $"${product.UnitPrice}",
                                  $"${product.TotalPrice}");
            }

            productsTable.ColumnAdjustment = ColumnAdjustment.AutoFitToWindow;
        }

        private void AddProductHeaders(Aspose.Pdf.Table table)
        {
            Aspose.Pdf.Row headerRow = table.Rows.Add();
            headerRow.FixedRowHeight = 20;

            headerRow.Border = new Aspose.Pdf.BorderInfo(Aspose.Pdf.BorderSide.Top | Aspose.Pdf.BorderSide.Bottom, 1f, Aspose.Pdf.Color.LightBlue);

            headerRow.BackgroundColor = Aspose.Pdf.Color.Transparent;


            AddProductHeader(headerRow, "Description");
            AddProductHeader(headerRow, "Quantity");
            AddProductHeader(headerRow, "Unit Price");
            AddProductHeader(headerRow, "Total Price");
        }

        private void AddProductHeader(Aspose.Pdf.Row row, string headerText)
        {
            Aspose.Pdf.Cell cell = row.Cells.Add();
            TextFragment text = new TextFragment(headerText);
            text.TextState.FontSize = 12;
            text.TextState.FontStyle = FontStyles.Bold;
            cell.Paragraphs.Add(text);
        }

        private void AddProductDetails(Aspose.Pdf.Table table, string description, string quantity, string unitPrice,
                                       string totalPrice)
        {
            Aspose.Pdf.Row row = table.Rows.Add();
            AddProductDetail(row, description);
            AddProductDetail(row, quantity);
            AddProductDetail(row, unitPrice);
            AddProductDetail(row, totalPrice);

        }

        private void AddProductDetail(Aspose.Pdf.Row row, string value)
        {
            Aspose.Pdf.Cell cell = row.Cells.Add();
            TextFragment text = new TextFragment(value);
            text.TextState.FontSize = 12;
            text.Margin.Top = 15;
            text.Margin.Bottom = 5;
            cell.Paragraphs.Add(text);
        }


        private void createUserInfo(Aspose.Pdf.Page page, Invoice invoice)
        {
            Aspose.Pdf.Table customerTable = new Aspose.Pdf.Table();
            customerTable.ColumnWidths = "300 300 300";
            customerTable.Margin.Bottom = 30;

            page.Paragraphs.Add(customerTable);



            Aspose.Pdf.Cell customerDetailsCell = customerTable.Rows.Add().Cells.Add();
            TextFragment nameAddressTitle = new TextFragment("Bill To");
            nameAddressTitle.Margin.Bottom = 5;
            nameAddressTitle.Margin.Right = 15;

            nameAddressTitle.TextState.FontSize = 16;
            nameAddressTitle.TextState.FontStyle = FontStyles.Bold;
            nameAddressTitle.HorizontalAlignment = HorizontalAlignment.Left;
            customerDetailsCell.Paragraphs.Add(nameAddressTitle);

            AddCustomerDetails(customerDetailsCell, "Name    :", invoice.CustomerName);
            AddCustomerDetails(customerDetailsCell, "Address :", invoice.BillingAddress);

            Aspose.Pdf.Cell contactDetailsCell = customerTable.Rows[0].Cells.Add();

            TextFragment contactTitle = new TextFragment("Contacts");
            contactTitle.Margin.Bottom = 5;
            contactTitle.Margin.Right = 15;

            contactTitle.TextState.FontSize = 16;
            contactTitle.TextState.FontStyle = FontStyles.Bold;
            contactTitle.HorizontalAlignment = HorizontalAlignment.Left;
            contactDetailsCell.Paragraphs.Add(contactTitle);

            AddCustomerDetails(contactDetailsCell, "Contact :", invoice.ContactNumber.ToString());
            AddCustomerDetails(contactDetailsCell, "Email   :", invoice.Email);



            Aspose.Pdf.Cell invoiceDetailsCell = customerTable.Rows[0].Cells.Add();
            TextFragment invoiceDetailTitle = new TextFragment("Invoice deails");
            invoiceDetailTitle.Margin.Bottom = 5;
            invoiceDetailTitle.Margin.Right = 15;

            invoiceDetailTitle.TextState.FontSize = 16;
            invoiceDetailTitle.TextState.FontStyle = FontStyles.Bold;
            invoiceDetailTitle.HorizontalAlignment = HorizontalAlignment.Left;
            invoiceDetailsCell.Paragraphs.Add(invoiceDetailTitle);

            AddCustomerDetails(invoiceDetailsCell, "Invoice Number  :", invoice.InvoiceNumber.ToString());
            AddCustomerDetails(invoiceDetailsCell, "Invoice Date    :", invoice.InvoiceDate.ToShortDateString());

            customerTable.ColumnAdjustment = ColumnAdjustment.AutoFitToWindow;

        }

        private void createFooter(Aspose.Pdf.Page page)
        {
            TextFragment footer = new TextFragment("Thank you for your business!");
            footer.TextState.FontSize = 12;
            footer.HorizontalAlignment = HorizontalAlignment.Center;
            double footerWidth = footer.TextState.FontSize * footer.Text.Length;
            double pageWidth = page.PageInfo.Width;
            double centerX = (pageWidth - footerWidth) / 2 - 50;

            footer.Position = new Position(centerX, 30);
            page.Paragraphs.Add(footer);
        }

        private void AddCustomerDetails(Aspose.Pdf.Cell cell, string label, string value)
        {
            TextFragment labelFragment = new TextFragment($"{label} {value}");
            labelFragment.TextState.FontSize = 12;
            labelFragment.Margin = new MarginInfo { Top = 5, Bottom = 5 };
            cell.Paragraphs.Add(labelFragment);

        }
        private void AddInvoiceDetails(Aspose.Pdf.Table table, string label, string value)
        {
            Aspose.Pdf.Row row = table.Rows.Add();
            Aspose.Pdf.Cell cell1 = row.Cells.Add();
            TextFragment labelFragment = new TextFragment(label);
            labelFragment.TextState.FontSize = 12;
            cell1.Paragraphs.Add(labelFragment);

            Aspose.Pdf.Cell cell2 = row.Cells.Add();
            TextFragment valueFragment = new TextFragment(value);
            valueFragment.TextState.FontSize = 12;
            cell2.Paragraphs.Add(valueFragment);
        }

        public void createHeader(Aspose.Pdf.Page page)
        {
            Aspose.Pdf.Table headerTable = new Aspose.Pdf.Table();
            headerTable.ColumnWidths = "150 300";
            headerTable.Margin.Top = 30;
            headerTable.Margin.Bottom = 40;
            page.Paragraphs.Add(headerTable);

            Aspose.Pdf.Row row1 = headerTable.Rows.Add();

            Aspose.Pdf.Cell logoCell = row1.Cells.Add();
            Image logo = new Image();
            logo.File = @"C:\Users\i44375\source\repos\FileConversion\Utils\image\invoice.png";
            logo.FixWidth = 50;
            logo.FixHeight = 50;
            logo.HorizontalAlignment = HorizontalAlignment.Left;
            logoCell.Paragraphs.Add(logo);

            Aspose.Pdf.Cell titleCell = row1.Cells.Add();
            TextFragment title = new TextFragment("YURGEN DEVS");
            title.TextState.FontSize = 26;
            title.TextState.FontStyle = FontStyles.Bold;
            title.HorizontalAlignment = HorizontalAlignment.Center;
            title.TextState.ForegroundColor = Color.CadetBlue;
            titleCell.Paragraphs.Add(title);

            Aspose.Pdf.Row row2 = headerTable.Rows.Add();

            Aspose.Pdf.Cell emptyCell = row2.Cells.Add();

            Aspose.Pdf.Cell creatorCell = row2.Cells.Add();
            TextFragment creatorText = new TextFragment("Crafting IT Solutions to Fit Your Every Need");
            creatorText.TextState.FontSize = 12;
            creatorText.TextState.FontStyle = FontStyles.Italic;
            creatorText.TextState.ForegroundColor = Color.LightGray;
            creatorText.HorizontalAlignment = HorizontalAlignment.Center;
            creatorCell.Paragraphs.Add(creatorText);

        }

    }
}
