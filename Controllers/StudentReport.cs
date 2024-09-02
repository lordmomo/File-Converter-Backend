using Aspose.Cells;
using Aspose.Pdf;
using Aspose.Pdf.Operators;
using Aspose.Pdf.Text;
using FileConversion.Entity;
using FileConversion.Service.ServiceInterface;
using FileConversion.Utils.Constants;
using Microsoft.AspNetCore.Mvc;
using System.Data;
using System.IO.Compression;
using System.Text;
using System.Xml;
using System.Xml.Xsl;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace FileConversion.Controllers
{
    public class StudentReport : Controller
    {
        private readonly StudentReportInterfae _studentReportInterfae;
        public StudentReport(StudentReportInterfae studentReportInterfae) {
            this._studentReportInterfae = studentReportInterfae;
        }
        
        [HttpPost("students/generate-marksheet")]
        public IActionResult UploadExcel([FromForm]List<IFormFile> files)
        {

            var message = _studentReportInterfae.CreateStudentReport(files);
            return Json(message);
            
        }

        //[HttpGet("students/download-pdfs")]
        //public IActionResult DownloadPdfs([FromQuery] List<string> pdfPaths)
        //{
        //    // Validate and process the list of PDF paths
        //    try
        //    {
        //        // Create a memory stream to hold the zip file
        //        using (var memoryStream = new MemoryStream())
        //        {
        //            using (var archive = new ZipArchive(memoryStream, ZipArchiveMode.Create, true))
        //            {
        //                foreach (var pdfPath in pdfPaths)
        //                {
        //                    var fileInfo = new FileInfo(pdfPath);
        //                    var entry = archive.CreateEntry(fileInfo.Name);

        //                    // Open the PDF file and write its contents to the entry in the zip archive
        //                    using (var fileStream = new FileStream(pdfPath, FileMode.Open, FileAccess.Read))
        //                    using (var entryStream = entry.Open())
        //                    {
        //                        fileStream.CopyTo(entryStream);
        //                    }
        //                }
        //            }

        //            // Return the zip file as a byte array
        //            return File(memoryStream.ToArray(), "application/zip", "StudentReports.zip");
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        // Handle exceptions and return appropriate response
        //        return StatusCode(500, new ResponseMessage { Success = false, Message = $"Failed to download PDFs: {ex.Message}" });
        //    }
        //}

        [HttpGet("students/download-single-pdf")]
        public IActionResult DownloadSinglePdf([FromQuery] string pdfPath)
        {
            var filePath =  pdfPath;

            if (!System.IO.File.Exists(filePath))
            {
                return NotFound();
            }

            try
            {
                var fileBytes = System.IO.File.ReadAllBytes(filePath);
                return File(fileBytes, "application/pdf", Path.GetFileName(filePath));
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { message = $"Failed to download PDF: {ex.Message}" });
            }
        }


       

        [HttpPost("students/table-merge")]
        public IActionResult UploadMErgeExcell()
        {
            AddTable_RowColSpan();
            return Json("success");
        }

        public static void AddTable_RowColSpan()
        {
            // Load source PDF document
            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document();
            pdfDocument.Pages.Add();

            // Initializes a new instance of the Table
            Aspose.Pdf.Table table = new Aspose.Pdf.Table
            {
                // Set the table border color as LightGray
                Border = new Aspose.Pdf.BorderInfo(Aspose.Pdf.BorderSide.All, .5f, Color.Black),
                // Set the border for table cells
                DefaultCellBorder = new Aspose.Pdf.BorderInfo(Aspose.Pdf.BorderSide.All, .5f, Color.Black)
            };

            // Add 1st row to table
            Aspose.Pdf.Row row1 = table.Rows.Add();
            for (int cellCount = 1; cellCount < 5; cellCount++)
            {
                // Add table cells
                row1.Cells.Add($"Test 1 {cellCount}");
            }

            // Add 2nd row to table
            Aspose.Pdf.Row row2 = table.Rows.Add();
            row2.Cells.Add($"Test 2 1");
            var cell = row2.Cells.Add($"Test 2 2");
            cell.ColSpan = 2;
            row2.Cells.Add($"Test 2 4");

            // Add 3rd row to table
            Aspose.Pdf.Row row3 = table.Rows.Add();
            row3.Cells.Add("Test 3 1");
            row3.Cells.Add("Test 3 2");
            row3.Cells.Add("Test 3 3");
            row3.Cells.Add("Test 3 4");

            // Add 4th row to table
            Aspose.Pdf.Row row4 = table.Rows.Add();
            row4.Cells.Add("Test 4 1");
            cell = row4.Cells.Add("Test 4 2");
            cell.RowSpan = 2;
            row4.Cells.Add("Test 4 3");
            row4.Cells.Add("Test 4 4");


            // Add 5th row to table
            row4 = table.Rows.Add();
            row4.Cells.Add("Test 5 1");
            row4.Cells.Add("Test 5 3");
            row4.Cells.Add("Test 5 4");

            // Add table object to first page of input document
            pdfDocument.Pages[1].Paragraphs.Add(table);

            // Save updated document containing table object
            pdfDocument.Save(Path.Combine(@"C:\Users\i44375\source\repos\FileConversion\Utils\FileConverter\TableMergeDemo\", "document_with_table_out.pdf"));
        }


        [HttpPost("students/upload-excel/template")]
        public IActionResult UploadExcelWithTemplate(List<IFormFile> files)
        {

            var messages = new List<string>();
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
                        ConvertExcelToPDFTemplateReports(filePath, @"C:\Users\i44375\source\repos\FileConversion\Utils\FileConverter\TemplatedReports");
                    }
                    else
                    {
                        messages.Add($"File '{file.FileName}' could not be converted. Only Excel files (.xls, .xlsx) are supported.");
                    }

                    System.IO.File.Delete(filePath);
                }

            }

            return Json("Succes");
        }

        private void ConvertExcelToPDFTemplateReports(string inputFilePath, string outputDirectory)
        {

            string templatePath = @"C:\\Users\\i44375\\source\\repos\\FileConversion\\Utils\\Template\\StudentMarksheetTemplate.docx";

            Workbook workbook = new Workbook(inputFilePath);

            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                for (int i = 1; i <= worksheet.Cells.MaxDataRow; i++)
                {
                    // Extract student data
                    string studentId = worksheet.Cells[i, 0].StringValue;
                    string studentName = worksheet.Cells[i, 1].StringValue;
                    string address = worksheet.Cells[i, 2].StringValue;
                    string grade = worksheet.Cells[i, 3].StringValue;
                    string contact = worksheet.Cells[i, 4].StringValue;
                    string fatherName = worksheet.Cells[i, 5].StringValue;
                    string motherName = worksheet.Cells[i, 6].StringValue;
                    int maths = Convert.ToInt32(worksheet.Cells[i, 7].DoubleValue);
                    int english = Convert.ToInt32(worksheet.Cells[i, 8].DoubleValue);
                    int nepali = Convert.ToInt32(worksheet.Cells[i, 9].DoubleValue);
                    int computer = Convert.ToInt32(worksheet.Cells[i, 10].DoubleValue);


                    int totalMarksObtained = maths + english + nepali + computer;

                    double percentage = (totalMarksObtained / 400.0) * 100;

                    string result;
                    if(percentage >= 80)
                    {
                        result = " Distinction";
                    }
                    else if(percentage >= 70 && percentage < 80){
                        result = "First divison";
                    }
                    else if (percentage >= 60 && percentage < 70)
                    {
                        result = "Second divison";
                    }
                    else if (percentage >= 40 && percentage < 60)
                    {
                        result = "Third divison";
                    }
                    else
                    {
                        result = "Fail";
                    }

                    Aspose.Words.Document doc = new Aspose.Words.Document(templatePath);

                    ReplacePlaceholder(doc, "StudentId", studentId);
                    ReplacePlaceholder(doc, "Name", studentName);
                    ReplacePlaceholder(doc, "Address", address);
                    ReplacePlaceholder(doc, "Grade", grade);
                    ReplacePlaceholder(doc, "Contact", contact);
                    ReplacePlaceholder(doc, "FathersName", fatherName);
                    ReplacePlaceholder(doc, "MothersName", motherName);
                    ReplacePlaceholder(doc, "Maths",  maths.ToString());
                    ReplacePlaceholder(doc, "English", english.ToString());
                    ReplacePlaceholder(doc, "Nepali", nepali.ToString());
                    ReplacePlaceholder(doc, "Computer", computer.ToString());
                    ReplacePlaceholder(doc, "TotalMarksObtained",totalMarksObtained.ToString());
                    ReplacePlaceholder(doc, "Percentage", percentage.ToString("0.00"));
                    ReplacePlaceholder(doc, "Result", result);

                    string outputPdfPath = Path.Combine(outputDirectory, $"{studentName.Replace(" ", "_")}_Report.pdf");
                    doc.Save(outputPdfPath);
                }
            }
        }

        private static void ReplacePlaceholder(Aspose.Words.Document doc, string placeholder, string value)
        {
            doc.Range.Replace("{" + placeholder + "}", value);
        }

        private void ConvertExcelToPDFReports(string inputFilePath, string outputDirectory)
        {
            Workbook workbook = new Workbook(inputFilePath);

            foreach (Worksheet worksheet in workbook.Worksheets)
            {


                for (int i = 1; i <= worksheet.Cells.MaxDataRow; i++)
                {
                    // Extract student data
                    string studentId = worksheet.Cells[i, 0].StringValue;
                    string studentName = worksheet.Cells[i, 1].StringValue;
                    string address = worksheet.Cells[i, 2].StringValue;
                    string grade = worksheet.Cells[i, 3].StringValue;
                    string contact = worksheet.Cells[i, 4].StringValue;
                    string fatherName = worksheet.Cells[i, 5].StringValue;
                    string motherName = worksheet.Cells[i, 6].StringValue;
                    int maths = Convert.ToInt32(worksheet.Cells[i, 7].DoubleValue); 
                    int english = Convert.ToInt32(worksheet.Cells[i, 8].DoubleValue);
                    int nepali = Convert.ToInt32(worksheet.Cells[i, 9].DoubleValue);
                    int computer = Convert.ToInt32(worksheet.Cells[i, 10].DoubleValue);

                    // Create a PDF document
                    Document pdfDocument = new Document();
                    Page pdfPage = pdfDocument.Pages.Add();

                    // Add title
                    TextFragment title = new TextFragment($"Annual Report of {studentName}, child of {fatherName} and {motherName}, who studies in grade {grade} ");
                    title.TextState.Font = FontRepository.FindFont("Arial");
                    title.TextState.FontSize = 14;
                    title.TextState.FontStyle = FontStyles.Bold;
                    title.Position = new Position(100, 700);
                    pdfPage.Paragraphs.Add(title);

                    Table detailsTable = new Table();
                    detailsTable.ColumnWidths = "150 300";
                    detailsTable.Top = 350;

                    detailsTable.Border = new BorderInfo(BorderSide.All, 0.5f);
                    detailsTable.DefaultCellPadding = new MarginInfo(4, 4, 4, 4);

                    AddCell(detailsTable, "Student ID:", studentId);
                    AddCell(detailsTable, "Student Name:", studentName);
                    AddCell(detailsTable, "Address:", address);
                    AddCell(detailsTable, "Grade:", grade);
                    AddCell(detailsTable, "Contact:", contact);
                    AddCell(detailsTable, "Father's Name:", fatherName);
                    AddCell(detailsTable, "Mother's Name:", motherName);

                    pdfPage.Paragraphs.Add(detailsTable);

                    Table marksTable = new Table();
                    marksTable.ColumnWidths = "200 100";
                    marksTable.Top = 200;

                    marksTable.Border = new BorderInfo(BorderSide.All, 0.5f);
                    marksTable.DefaultCellPadding = new MarginInfo(4, 4, 4, 4);

                    marksTable.Rows.Add();
                    Aspose.Pdf.Cell headerCell1 = marksTable.Rows[0].Cells.Add("Subject");
                    headerCell1.BackgroundColor = Color.FromRgb(System.Drawing.Color.LightGray);
                    headerCell1.DefaultCellTextState.FontStyle = FontStyles.Bold;
                    headerCell1.Border = new BorderInfo(BorderSide.All, 0.5f);

                    
                    Aspose.Pdf.Cell headerCell2 = marksTable.Rows[0].Cells.Add("Marks Obtained");
                    headerCell2.BackgroundColor = Color.FromRgb(System.Drawing.Color.LightGray);
                    headerCell2.DefaultCellTextState.FontStyle = FontStyles.Bold;
                    headerCell2.Border = new BorderInfo(BorderSide.All, 0.5f);

                    AddRowToTable(marksTable, "Maths", maths.ToString());
                    AddRowToTable(marksTable, "English", english.ToString());
                    AddRowToTable(marksTable, "Nepali", nepali.ToString());
                    AddRowToTable(marksTable, "Computer", computer.ToString());

                    pdfPage.Paragraphs.Add(marksTable);

                    string outputPdfPath = Path.Combine(outputDirectory, $"{studentName.Replace(" ", "_")}_Report.pdf");
                    pdfDocument.Save(outputPdfPath);
                }
            }
        }



        [HttpPost("students/upload-excel/xml/template")]

        public IActionResult TempalteViewXml(List<IFormFile> files)
        {
            var messages = new List<string>();

            var xsltPath = "C:\\Users\\i44375\\source\\repos\\FileConversion\\Utils\\Template\\StudentMarkSheet.xslt";

            var xmlPath = "C:\\Users\\i44375\\source\\repos\\FileConversion\\Utils\\Template\\StudentTemplateXml.xml";

            //var outputPath = "C:\\Users\\i44375\\source\\repos\\FileConversion\\Utils\\FileConverter\\XmlReports\\";

            Aspose.Pdf.Document pdf = new Aspose.Pdf.Document();


            try
            {
                var xmlDocString = GenerateXmlFromExcel(files);

                var xmlDoc = new XmlDocument();
                xmlDoc.LoadXml(xmlDocString);

                pdf.BindXml(xmlPath, xsltPath);
                foreach (XmlElement studentElement in xmlDoc.DocumentElement.SelectNodes("StudentInformation"))
                {
                    var studentId = studentElement.SelectSingleNode("StudentID").InnerText;
                    var studentName = studentElement.SelectSingleNode("Name").InnerText;
                    var address = studentElement.SelectSingleNode("Address").InnerText;
                    var grade = studentElement.SelectSingleNode("Grade").InnerText;
                    var contact = studentElement.SelectSingleNode("Contact").InnerText;
                    var fathersName = studentElement.SelectSingleNode("FathersName").InnerText;
                    var mothersName = studentElement.SelectSingleNode("MothersName").InnerText;
                    var maths = studentElement.SelectSingleNode("Maths").InnerText;
                    var english = studentElement.SelectSingleNode("English").InnerText;
                    var nepali = studentElement.SelectSingleNode("Nepali").InnerText;
                    var computer = studentElement.SelectSingleNode("Computer").InnerText;
                    var totalMarksObtained = studentElement.SelectSingleNode("TotalMarksObtained").InnerText;
                    var percentage = studentElement.SelectSingleNode("Percentage").InnerText;
                    var result = studentElement.SelectSingleNode("Result").InnerText;


                    var page = pdf.Pages.Add();
                    var textFragment = new TextFragment($"Student ID: {studentId}, Name: {studentName}");
                    page.Paragraphs.Add(textFragment);

                    // Save PDF to a specific folder
                    var outputPath = $"C:\\Users\\i44375\\source\\repos\\FileConversion\\Utils\\FileConverter\\XmlReports\\{studentName}_{studentId}.pdf";
                    pdf.Save(outputPath);

                    messages.Add($"PDF generated for Student ID: {studentId}, Name: {studentName}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                throw;
            }


       
            return Json("Succes");

        }

        private string GenerateXmlFromExcel(List<IFormFile> files)
        {
            var xmlDoc = new XmlDocument();
            var root = xmlDoc.CreateElement("StudentMarkSheet");
            xmlDoc.AppendChild(root);

            foreach (var file in files)
            {
                if (file.Length > 0)
                {
                    using (var stream = new MemoryStream())
                    {
                        file.CopyTo(stream);
                        stream.Position = 0;

                        Workbook workbook = new Workbook(stream);
                        //Worksheet worksheet = workbook.Worksheets[0];

                        foreach( Worksheet worksheet in workbook.Worksheets )
                        {
                            for (int i = 1; i <= worksheet.Cells.MaxDataRow; i++)
                            {


                                string studentId = worksheet.Cells[i, 0].StringValue;
                                string studentName = worksheet.Cells[i, 1].StringValue;
                                string address = worksheet.Cells[i, 2].StringValue;
                                string grade = worksheet.Cells[i, 3].StringValue;
                                string contact = worksheet.Cells[i, 4].StringValue;
                                string fatherName = worksheet.Cells[i, 5].StringValue;
                                string motherName = worksheet.Cells[i, 6].StringValue;
                                int maths = Convert.ToInt32(worksheet.Cells[i, 7].DoubleValue);
                                int english = Convert.ToInt32(worksheet.Cells[i, 8].DoubleValue);
                                int nepali = Convert.ToInt32(worksheet.Cells[i, 9].DoubleValue);
                                int computer = Convert.ToInt32(worksheet.Cells[i, 10].DoubleValue);

                                int totalMarksObtained = maths + english + nepali + computer;
                                double percentage = (totalMarksObtained / 400.0) * 100;

                                string result;
                                if (percentage >= 80)
                                    result = "Distinction";
                                else if (percentage >= 70 && percentage < 80)
                                    result = "First Division";
                                else if (percentage >= 60 && percentage < 70)
                                    result = "Second Division";
                                else if (percentage >= 40 && percentage < 60)
                                    result = "Third Division";
                                else
                                    result = "Fail";

                                var studentInformation = xmlDoc.CreateElement("StudentInformation");
                                studentInformation.AppendChild(CreateElement(xmlDoc, "StudentID", studentId));
                                studentInformation.AppendChild(CreateElement(xmlDoc, "Name", studentName));
                                studentInformation.AppendChild(CreateElement(xmlDoc, "Address", address));
                                studentInformation.AppendChild(CreateElement(xmlDoc, "Grade", grade));
                                studentInformation.AppendChild(CreateElement(xmlDoc, "Contact", contact));
                                studentInformation.AppendChild(CreateElement(xmlDoc, "FatherName", fatherName));
                                studentInformation.AppendChild(CreateElement(xmlDoc, "MotherName", motherName));
                                root.AppendChild(studentInformation);

                                var academicDetails = xmlDoc.CreateElement("AcademicDetails");
                                academicDetails.AppendChild(CreateSubjectElement(xmlDoc, "Maths", maths));
                                academicDetails.AppendChild(CreateSubjectElement(xmlDoc, "English", english));
                                academicDetails.AppendChild(CreateSubjectElement(xmlDoc, "Nepali", nepali));
                                academicDetails.AppendChild(CreateSubjectElement(xmlDoc, "Computer", computer));
                                root.AppendChild(academicDetails);

                                var overallResult = xmlDoc.CreateElement("OverallResult");
                                overallResult.AppendChild(CreateElement(xmlDoc, "TotalMarksObtained", totalMarksObtained.ToString()));
                                overallResult.AppendChild(CreateElement(xmlDoc, "Percentage", percentage.ToString()));
                                overallResult.AppendChild(CreateElement(xmlDoc, "Result", result));
                                root.AppendChild(overallResult);
                            }
                        }
                        
                    }
                }
            }

            return xmlDoc.OuterXml;
        }

        private XmlElement CreateElement(XmlDocument xmlDoc, string name, string value)
        {
            var element = xmlDoc.CreateElement(name);
            element.InnerText = value;
            return element;
        }

        private XmlElement CreateSubjectElement(XmlDocument xmlDoc, string subject, int marks)
        {
            var subjectMarks = xmlDoc.CreateElement("SubjectMarks");
            subjectMarks.AppendChild(CreateElement(xmlDoc, "Subject", subject));
            subjectMarks.AppendChild(CreateElement(xmlDoc, "MarksObtained", marks.ToString()));
            subjectMarks.AppendChild(CreateElement(xmlDoc, "TotalMarks", "100"));
            return subjectMarks;
        }



        private MemoryStream TransformXmltoHtml(string inputXml, string xsltString)
        {
            

            try
            {
                XslCompiledTransform transform = new XslCompiledTransform();
                using (XmlReader reader = XmlReader.Create(new StringReader(xsltString)))
                {
                    transform.Load(reader);
                }

                MemoryStream outputStream = new MemoryStream();

                using (StringReader sReader = new StringReader(inputXml))
                {
                    using (XmlReader reader = XmlReader.Create(sReader))
                    {
                        XmlWriterSettings settings = new XmlWriterSettings();
                        settings.OmitXmlDeclaration = true;
                        settings.ConformanceLevel = ConformanceLevel.Document;

                        using (XmlWriter xmlWriter = XmlWriter.Create(outputStream, settings))
                        {
                            transform.Transform(reader, xmlWriter);
                        }
                    }
                }

                outputStream.Position = 0;
                return outputStream;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error transforming XML: {ex.Message}");
                throw;
             }
        }

        private void SaveHtmlAsPdf(MemoryStream htmlStream, string outputPdfPath)
        {
            var options = new Aspose.Pdf.HtmlLoadOptions();
            options.PageInfo.Height = 595; // Set page height (in points, 1 inch = 72 points)
            options.PageInfo.Width = 420; // Set page width (in points, 1 inch = 72 points)

            var storeDir = "C:\\Users\\i44375\\source\\repos\\FileConversion\\Utils\\FileConverter\\XmlTemplate\\";

            var pdfDocument = new Aspose.Pdf.Document(htmlStream, options);
            pdfDocument.Save(storeDir);
        }
        /////

        private void ConvertXmlToPDFTemplateReports(string inputFilePath, string outputDirectory)
        {
            Workbook workbook = new Workbook(inputFilePath);
            var xmlTemplatePath = "C:\\Users\\i44375\\source\\repos\\FileConversion\\Utils\\Template\\StudentTemplateXml.xml";
            string xmlTemplate = System.IO.File.ReadAllText(xmlTemplatePath);


            foreach (Worksheet worksheet in workbook.Worksheets)
            {


                for (int i = 1; i <= worksheet.Cells.MaxDataRow; i++)
                {

                    GenerateXmlFromRow(xmlTemplate, worksheet.Cells, i, outputDirectory);

                }
            }
        }

        static void GenerateXmlFromRow(string xmlTemplate, Aspose.Cells.Cells cells, int rowIndex, string outputDirectory)
        {
            var storeDir = outputDirectory;

            string studentId = cells[rowIndex, 0].StringValue;
            string studentName = cells[rowIndex, 1].StringValue;
            string address = cells[rowIndex, 2].StringValue;
            string grade = cells[rowIndex, 3].StringValue;
            string contact = cells[rowIndex, 4].StringValue;
            string fatherName = cells[rowIndex, 5].StringValue;
            string motherName = cells[rowIndex, 6].StringValue;
            int maths = cells[rowIndex, 7].IntValue;
            int english = cells[rowIndex, 8].IntValue;
            int nepali = cells[rowIndex, 9].IntValue;
            int computer = cells[rowIndex, 10].IntValue;

            int totalMarksObtained = (maths + english + nepali + computer);

            double percentage = (totalMarksObtained / 400.0) * 100;

            string result;
            if (percentage >= 80)
            {
                result = " Distinction";
            }
            else if (percentage >= 70 && percentage < 80)
            {
                result = "First divison";
            }
            else if (percentage >= 60 && percentage < 70)
            {
                result = "Second divison";
            }
            else if (percentage >= 40 && percentage < 60)
            {
                result = "Third divison";
            }
            else
            {
                result = "Fail";
            }

            string xmlData = xmlTemplate.Replace("{StudentId}", studentId)
                                        .Replace("{Name}", studentName)
                                        .Replace("{Address}", address)
                                        .Replace("{Grade}", grade)
                                        .Replace("{Contact}", contact)
                                        .Replace("{FatherName}", fatherName)
                                        .Replace("{MotehrName}", motherName)
                                        .Replace("{Maths}", maths.ToString())
                                        .Replace("{English}", english.ToString())
                                        .Replace("{Nepali}", nepali.ToString())
                                        .Replace("{Computer}", computer.ToString())
                                        .Replace("{TotalMarksObtained}", totalMarksObtained.ToString())
                                        .Replace("{Percentage}", percentage.ToString())
                                        .Replace("{Result}", result);

            var outputFileName = $"output_{studentName}.pdf";

            SaveXmlAsPdf(xmlData, Path.Combine(storeDir, outputFileName));

        }
        static void SaveXmlAsPdf(string xmlData, string outputFilePath)
        {


            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document();

            Page page = pdfDocument.Pages.Add();
            const double PageHeight = 800; 
            const double Margin = 50;

            TextBuilder textBuilder = new TextBuilder(page);


            double contentHeight = 0;

            TextFragment textFragment = new TextFragment(xmlData);

            textBuilder.AppendText(textFragment);

            pdfDocument.Save(outputFilePath);

        }
   

        private void AddCell(Table table, string label, string value)
        {
            Aspose.Pdf.Row row = table.Rows.Add();
            Aspose.Pdf.Cell cell1 = row.Cells.Add(label);
            cell1.DefaultCellTextState.FontStyle = FontStyles.Bold;
            cell1.Border = new BorderInfo(BorderSide.All, 0.5f);

            Aspose.Pdf.Cell cell2 = row.Cells.Add(value);
            cell2.Border = new BorderInfo(BorderSide.All, 0.5f);
        }

        private void AddRowToTable(Table table, string subject, string marks)
        {
            Aspose.Pdf.Row row = table.Rows.Add();
            Aspose.Pdf.Cell cell1 = row.Cells.Add(subject);
            cell1.Border = new BorderInfo(BorderSide.All, 0.5f);

            Aspose.Pdf.Cell cell2 = row.Cells.Add(marks);
            cell2.Border = new BorderInfo(BorderSide.All, 0.5f);
        }
       
    }


}
