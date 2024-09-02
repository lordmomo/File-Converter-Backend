using Aspose.Cells;
using Aspose.Pdf;
using Aspose.Pdf.Text;
using Aspose.Words.Tables;
using FileConversion.Entity;
using FileConversion.Service.ServiceInterface;
using FileConversion.Utils.Constants;
using System.IO;
using System.Runtime.Intrinsics.Arm;

namespace FileConversion.Service.ServiceImplementation
{
    public class StudentReportImplementation : StudentReportInterfae
    {
        List<string> previewPathForFiles = new List<string>();
        public ReportGenerationResponseMessage CreateStudentReport(List<IFormFile> files)
        {
            var studentList = new List<Student>();
            var generatedPdfPaths = new List<string>();

            var previewReportPath = new List<string>();

            try
            {


                if (files == null || files.Count == 0)
                {
                    return new ReportGenerationResponseMessage { Success = false, Message = Message.fileNotProvided };
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
                            studentList.Clear();
                            ExtractStudentDataFromFiles(filePath, studentList);
                        }
                        else
                        {
                            return new ReportGenerationResponseMessage { Success = false, Message = Message.onlyExcelFilesAllowed };

                        }
                        if (studentList.Count > 0)
                        {
                            SaveStudentReportInPdf(studentList, Constants.studentReportStoragePath);
                        }
                        string previewPdfPath = Path.Combine(Constants.studentReportStoragePath, $"Student_{studentList[0].StudentId}_Report.pdf");
                        previewReportPath.Add(previewPdfPath);
                        System.IO.File.Delete(filePath);
                    }

                    foreach (var student in studentList)
                    {
                        string outputPdfPath = Path.Combine(Constants.studentReportStoragePath, $"Student_{student.StudentId}_Report.pdf");
                        generatedPdfPaths.Add(outputPdfPath);
                    }
                }


                setPreviewPath(previewReportPath);
                return new ReportGenerationResponseMessage { Success = true, Message = Message.successfullySavePdfs, Data = generatedPdfPaths, PreviewData = previewReportPath };

            }
            catch (ArgumentOutOfRangeException ex)
            {
                Console.WriteLine("Index error: ",ex.Message);
                return new ReportGenerationResponseMessage { Success = false, Message = Message.unmatchedInputFileForTemplate };
            }
            catch (Exception ex)
            {
                Console.WriteLine("An unexpected error occured:", ex.Message);
                return new ReportGenerationResponseMessage { Success = false, Message = Message.unexpectedErrorWhileProcessingFiles };
            }

        }

        private void setPreviewPath(List<string> previewReportPath)
        {
            previewPathForFiles = previewReportPath;
        }

        //private List<string> getPreviewPath()
        //{
        //    return previewPathForFiles ;
        //}

        public List<string> getPreviewPath()
        {
            return previewPathForFiles ;
        }

        private void ExtractStudentDataFromFiles(string inputFilePath, List<Student> studentList)
        {
            Workbook workbook = new Workbook(inputFilePath);
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                for (int i = 1; i <= worksheet.Cells.MaxDataRow; i++)
                {
                    try
                    {

                        string studentId = worksheet.Cells[i, 0].StringValue;
                        string studentName = worksheet.Cells[i, 1].StringValue;
                        string address = worksheet.Cells[i, 2].StringValue;
                        DateTime dateOfBirth = worksheet.Cells[i, 3].DateTimeValue;
                        int grade = worksheet.Cells[i, 4].IntValue;
                        double contact = worksheet.Cells[i, 5].DoubleValue;
                        string fatherName = worksheet.Cells[i, 6].StringValue;
                        string motherName = worksheet.Cells[i, 7].StringValue;

                        int mathTh = worksheet.Cells[i, 8].IntValue;
                        int mathPr = worksheet.Cells[i, 9].IntValue;
                        int englishTh = worksheet.Cells[i, 10].IntValue;
                        int englishPr = worksheet.Cells[i, 11].IntValue;
                        int nepaliTh = worksheet.Cells[i, 12].IntValue;
                        int nepaliPr = worksheet.Cells[i, 13].IntValue;
                        int computerTh = worksheet.Cells[i, 14].IntValue;
                        int computerPr = worksheet.Cells[i, 15].IntValue;
                        int scienceTh = worksheet.Cells[i, 16].IntValue;
                        int sciencePr = worksheet.Cells[i, 17].IntValue;
                        int socialStudiesTh = worksheet.Cells[i, 18].IntValue;
                        int socialStudiesPr = worksheet.Cells[i, 19].IntValue;
                        int hpeTh = worksheet.Cells[i, 20].IntValue;
                        int hpePr = worksheet.Cells[i, 21].IntValue;
                        int accountancyTh = worksheet.Cells[i, 22].IntValue;
                        int accountancyPr = worksheet.Cells[i, 23].IntValue;

                        int passMarks = worksheet.Cells[i, 24].IntValue;
                        int fullMarks = worksheet.Cells[i, 25].IntValue;
                        int totalMarks = worksheet.Cells[i, 26].IntValue;
                        double percentage = worksheet.Cells[i, 27].DoubleValue;
                        string remark = worksheet.Cells[i, 28].StringValue;


                        var studentReport = new Student
                        {
                            StudentId = studentId,
                            StudentName = studentName,
                            Address = address,
                            DateOfBirth = dateOfBirth,
                            Grade = grade,
                            Contact = contact,
                            FathersName = fatherName,
                            MothersName = motherName,
                            MathsTh = mathTh,
                            MathsPr = mathPr,
                            EnglishTh = englishTh,
                            EnglishPr = englishPr,
                            NepaliTh = nepaliTh,
                            NepaliPr = nepaliPr,
                            ComputerTh = computerTh,
                            ComputerPr = computerPr,
                            ScienceTh = scienceTh,
                            SciencePr = sciencePr,
                            SocailStudiesTh = socialStudiesTh,
                            SocailStudiesPr = socialStudiesPr,
                            HPETh = hpeTh,
                            HPEPr = hpePr,
                            AccountancyTh = accountancyTh,
                            AccountancyPr = accountancyPr,
                            PassMarks = passMarks,
                            FullMarks = fullMarks,
                            TotalMarks = totalMarks,
                            Percentage = percentage,
                            Remark = remark
                        };

                        studentList.Add(studentReport);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing row {i}: {ex.Message}");

                    }
                }
            }
        }

        private void SaveStudentReportInPdf(List<Student> studentList, string outputFilePath)
        {
            foreach (var student in studentList)
            {
                Document document = new Document();

                Aspose.Pdf.Page page = document.Pages.Add();
                page.SetPageSize(PageSize.A4.Width, PageSize.A4.Height);

                page.PageInfo.Margin = new MarginInfo { Top = 30, Bottom = 20, Left = 50, Right = 30 };

                createHeader(page);

                createUserInfo(page, student);

                createStudentResultTable(page, student);

                createResultRemarks(page, student);

                createCheckedByField(page);


                string outputPdfPath = Path.Combine(outputFilePath, "Student_" + (student.StudentId) + "_Report.pdf");

                document.Save(outputPdfPath);
            }
        }

        private void createCheckedByField(Page page)
        {

            Image logo = new Image();
            logo.File = Constants.checkedBySignaturePath;
            logo.FixWidth = 70;
            logo.FixHeight = 50;
            logo.HorizontalAlignment = HorizontalAlignment.Left;
            page.Paragraphs.Add(logo);



            TextFragment checkedByLabel = new TextFragment(Constants.checkedByTitle);
            checkedByLabel.TextState.FontSize = 12;
            checkedByLabel.HorizontalAlignment = HorizontalAlignment.Left;
            checkedByLabel.Margin.Top = 20;
            checkedByLabel.Margin.Bottom = 5;
            page.Paragraphs.Add(checkedByLabel);

            TextFragment checkedByName = new TextFragment(Constants.checkedByPersonName);
            checkedByName.TextState.FontSize = 12;
            checkedByName.TextState.FontStyle = FontStyles.Bold;
            checkedByName.HorizontalAlignment = HorizontalAlignment.Left;
            checkedByName.Margin.Top = 5;
            checkedByName.Margin.Bottom = 5;
            page.Paragraphs.Add(checkedByName);

            TextFragment dateOfIssue = new TextFragment($"DATE OF ISSUE : {DateTime.Now.Date.ToString("yyyy-MM-dd")}");
            dateOfIssue.TextState.FontSize = 12;
            dateOfIssue.HorizontalAlignment = HorizontalAlignment.Left;
            dateOfIssue.Margin.Top = 5;
            dateOfIssue.Margin.Bottom = 5;
            page.Paragraphs.Add(dateOfIssue);
        }

        private void createResultRemarks(Page page, Student student)
        {


            TextFragment grandTotalLabel = new TextFragment($"Grand Total     : {student.TotalMarks}");
            grandTotalLabel.TextState.FontSize = 12;
            grandTotalLabel.HorizontalAlignment = HorizontalAlignment.Right;
            grandTotalLabel.Margin.Top = 20;
            page.Paragraphs.Add(grandTotalLabel);

            TextFragment percentageAchieved = new TextFragment($"Percentage     : {student.Percentage * 100}%");
            percentageAchieved.TextState.FontSize = 12;
            percentageAchieved.HorizontalAlignment = HorizontalAlignment.Right;
            percentageAchieved.Margin.Top = 15;
            percentageAchieved.Margin.Bottom = 5;
            page.Paragraphs.Add(percentageAchieved);

            string calculatedRemark;

            if (student.Remark.Equals("Pass"))
            {
                double percentage = student.Percentage * 100;

                if (percentage >= 80)
                {
                    calculatedRemark = "Distinction";
                }
                else if (percentage >= 70 && percentage < 80)
                {
                    calculatedRemark = "First divison";
                }
                else if (percentage >= 60 && percentage < 70)
                {
                    calculatedRemark = "Second divison";
                }
                else if (percentage >= 40 && percentage < 60)
                {
                    calculatedRemark = "Third divison";
                }
                else
                {
                    calculatedRemark = "Fourth division";
                }
            }
            else
            {
                calculatedRemark = "Failed";
            }
            TextFragment remarkAchieved = new TextFragment($"Division     : {calculatedRemark}");
            remarkAchieved.TextState.FontSize = 12;
            remarkAchieved.HorizontalAlignment = HorizontalAlignment.Right;
            remarkAchieved.Margin.Top = 15;
            remarkAchieved.Margin.Bottom = 5;
            page.Paragraphs.Add(remarkAchieved);



        }

        public void createHeader(Aspose.Pdf.Page page)
        {
            Aspose.Pdf.Table headerTable = new Aspose.Pdf.Table();
            headerTable.ColumnWidths = "150 300";
            headerTable.Margin.Top = 30;
            headerTable.Margin.Bottom = 30;
            page.Paragraphs.Add(headerTable);

            Aspose.Pdf.Row row1 = headerTable.Rows.Add();

            Aspose.Pdf.Cell logoCell = row1.Cells.Add();
            Image logo = new Image();
            logo.File = Constants.schoolLogoPath;
            logo.FixWidth = 60;
            logo.FixHeight = 60;
            logo.HorizontalAlignment = HorizontalAlignment.Left;
            logoCell.Paragraphs.Add(logo);

            Aspose.Pdf.Cell titleCell = row1.Cells.Add();
            TextFragment title = new TextFragment(Constants.schoolName);
            title.TextState.FontSize = 26;
            title.TextState.FontStyle = FontStyles.Bold;
            title.HorizontalAlignment = HorizontalAlignment.Center;
            title.TextState.ForegroundColor = Color.CadetBlue;
            titleCell.Paragraphs.Add(title);

            Aspose.Pdf.Row row2 = headerTable.Rows.Add();

            Aspose.Pdf.Cell emptyCell = row2.Cells.Add();
                
            Aspose.Pdf.Cell creatorCell = row2.Cells.Add();
            TextFragment creatorText = new TextFragment(Constants.schoolSlogan);
            creatorText.TextState.FontSize = 12;
            creatorText.TextState.FontStyle = FontStyles.Italic;
            creatorText.TextState.ForegroundColor = Color.LightGray;
            creatorText.HorizontalAlignment = HorizontalAlignment.Center;
            creatorCell.Paragraphs.Add(creatorText);

            TextFragment gradeSheetLabel = new TextFragment(Constants.gradeSheetTitle);
            gradeSheetLabel.TextState.FontSize = 18;
            gradeSheetLabel.TextState.FontStyle = FontStyles.Bold;
            gradeSheetLabel.HorizontalAlignment = HorizontalAlignment.Center;
            gradeSheetLabel.VerticalAlignment = VerticalAlignment.Center;
            gradeSheetLabel.Margin.Bottom = 10;
            page.Paragraphs.Add(gradeSheetLabel);
        }


        private void createUserInfo(Aspose.Pdf.Page page, Student student)
        {
            TextFragment userText1 = new TextFragment($"The grade secured by {student.StudentName}");

            userText1.Margin.Top = 5;
            userText1.Margin.Bottom = 5;
            userText1.TextState.FontSize = 12;
            page.Paragraphs.Add(userText1);

            TextFragment userText2 = new TextFragment($"Date of Birth:  {student.DateOfBirth.Date.ToString("yyyy-MM-dd")} A.D.");

            userText2.Margin.Top = 5;
            userText2.Margin.Bottom = 5;
            userText2.TextState.FontSize = 12;

            page.Paragraphs.Add(userText2);

            TextFragment userText3 = new TextFragment($"Address: {student.Address}                              Symbol No.:  {student.StudentId}");

            userText3.Margin.Top = 5;
            userText3.Margin.Bottom = 5;
            userText3.TextState.FontSize = 12;

            page.Paragraphs.Add(userText3);

            TextFragment userText4 = new TextFragment($"In the annual examination of {DateTime.Now.Year} A.D. are given below");

            userText4.Margin.Top = 5;
            userText4.Margin.Bottom = 5;
            userText4.TextState.FontSize = 12;

            page.Paragraphs.Add(userText4);
        }

        private void createStudentResultTable(Aspose.Pdf.Page page, Student student)
        {
            Aspose.Pdf.Table studentTable = new Aspose.Pdf.Table();
            studentTable.ColumnWidths = "40 150 45 45 45 45 50 50 45";
            studentTable.Margin.Top = 10;

            //studentTable.Border = new Aspose.Pdf.BorderInfo(Aspose.Pdf.BorderSide.All, 0.1f, Aspose.Pdf.Color.Black);

            page.Paragraphs.Add(studentTable);

            AddMarksheetHeaders(studentTable);

            AddMarksRow(studentTable, "1", "English", "4", student.FullMarks.ToString(), student.PassMarks.ToString(), student.EnglishTh, student.EnglishPr);
            AddMarksRow(studentTable, "2", "Nepali", "4", student.FullMarks.ToString(), student.PassMarks.ToString(), student.NepaliTh, student.NepaliPr);
            AddMarksRow(studentTable, "3", "Maths", "4", student.FullMarks.ToString(), student.PassMarks.ToString(), student.MathsTh, student.MathsPr);
            AddMarksRow(studentTable, "4", "Computer", "4", student.FullMarks.ToString(), student.PassMarks.ToString(), student.ComputerTh, student.ComputerPr);
            AddMarksRow(studentTable, "5", "Science", "4", student.FullMarks.ToString(), student.PassMarks.ToString(), student.ScienceTh, student.SciencePr);
            AddMarksRow(studentTable, "6", "Social Studies", "4", student.FullMarks.ToString(), student.PassMarks.ToString(), student.SocailStudiesTh, student.SocailStudiesPr);
            AddMarksRow(studentTable, "7", "Health, Pop & Env", "4", student.FullMarks.ToString(), student.PassMarks.ToString(), student.HPETh, student.HPEPr);
            AddMarksRow(studentTable, "8", "Accountancy", "4", student.FullMarks.ToString(), student.PassMarks.ToString(), student.AccountancyTh, student.AccountancyPr);

            //studentTable.ColumnAdjustment = ColumnAdjustment.AutoFitToWindow;
        }

        private void AddMarksheetHeaders(Aspose.Pdf.Table table)
        {
            Aspose.Pdf.Row headerRow1 = table.Rows.Add();
            //headerRow1.FixedRowHeight = 30;
            headerRow1.Border = new Aspose.Pdf.BorderInfo(Aspose.Pdf.BorderSide.Box, 0.1f, Aspose.Pdf.Color.Black);
            headerRow1.BackgroundColor = Aspose.Pdf.Color.Transparent;
            headerRow1.FixedRowHeight = 40;

            //var cell1 =headerRow1.Cells.Add("Code");
            //cell1.RowSpan = 2;
            //var cell2 =headerRow1.Cells.Add("Subject");
            //cell2.RowSpan = 2;
            //var cell3 = headerRow1.Cells.Add("Credit Hour");
            //cell3.RowSpan = 2;
            //var cell4 = headerRow1.Cells.Add("Full Marks");
            //cell4.ColSpan = 2;
            //var cell5 = headerRow1.Cells.Add("Pass Marks");
            //cell5.RowSpan = 2;
            //var cell6 = headerRow1.Cells.Add("Marks Obtained (Th)");
            //cell6.RowSpan = 2;
            //var cell7 = headerRow1.Cells.Add("Marks Obtained (Pr)");
            //cell7.RowSpan = 2;
            //var cell8 = headerRow1.Cells.Add("Total Marks");
            //cell8.RowSpan = 2;

            SetMarksheetHeaders(headerRow1, "Code");
            SetMarksheetHeaders(headerRow1, "Subject");
            SetMarksheetHeaders(headerRow1, "Credit Hour");

            SetMarksheetHeaders(headerRow1, "Full Marks (Th)");
            SetMarksheetHeaders(headerRow1, "Full Marks (Pr)");


            //headerRow1.Cells.Add("FullMarks");
            //var cell3 = headerRow1.Cells.Add("Full Marks");
            //cell3.ColSpan = 2;

            //cell3.Border = new Aspose.Pdf.BorderInfo(Aspose.Pdf.BorderSide.All, 1f, Aspose.Pdf.Color.Red);
            //cell3.BackgroundColor = Aspose.Pdf.Color.Red;

            SetMarksheetHeaders(headerRow1, "Pass Marks");
            SetMarksheetHeaders(headerRow1, "Marks Obtained (Th)");
            SetMarksheetHeaders(headerRow1, "Marks Obtained (Pr)");
            SetMarksheetHeaders(headerRow1, "Total Marks ");


            //var cell4 = headerRow1.Cells.Add("Full Marks");
            ////cell3.ColSpan = 2;

            //cell4.Border = new Aspose.Pdf.BorderInfo(Aspose.Pdf.BorderSide.All, 1f, Aspose.Pdf.Color.Red);
            //cell4.BackgroundColor = Aspose.Pdf.Color.Red;



            //headerRow1.Cells[3].ColSpan = 2;


            //table.Rows[0].Cells[0].RowSpan = 2;
            //table.Rows[0].Cells[1].RowSpan = 2;
            //table.Rows[0].Cells[2].RowSpan = 2;
            ////table.Rows[0].Cells[3].RowSpan = 2;
            //table.Rows[0].Cells[4].RowSpan = 2;
            //table.Rows[0].Cells[5].RowSpan = 2;
            //table.Rows[0].Cells[6].RowSpan = 2;
            //table.Rows[0].Cells[7].RowSpan = 2;



            //Aspose.Pdf.Row headerRow2 = table.Rows.Add();
            ////headerRow2.FixedRowHeight = 20;

            //headerRow2.Border = new Aspose.Pdf.BorderInfo(Aspose.Pdf.BorderSide.All, 0.1f, Aspose.Pdf.Color.Black);

            //headerRow2.BackgroundColor = Aspose.Pdf.Color.Transparent;
            //headerRow2.FixedRowHeight = 80;


            //// Add sub-headers under "Full Marks"
            // var fullMarksTHCell = headerRow2.Cells.Add("TH");
            // fullMarksTHCell.Border = new Aspose.Pdf.BorderInfo(Aspose.Pdf.BorderSide.All, 1f, Aspose.Pdf.Color.Yellow);
            // var fullMarksPRCell = headerRow2.Cells.Add("PR");
            // fullMarksPRCell.Border = new Aspose.Pdf.BorderInfo(Aspose.Pdf.BorderSide.All, 1f, Aspose.Pdf.Color.Blue);    
            //// Adjust the cells' borders to ensure they align correctly
            // fullMarksTHCell.Border = new Aspose.Pdf.BorderInfo(Aspose.Pdf.BorderSide.All, 1f, Aspose.Pdf.Color.Yellow);
            // fullMarksPRCell.Border = new Aspose.Pdf.BorderInfo(Aspose.Pdf.BorderSide.All, 1f, Aspose.Pdf.Color.Blue);





            //SetMarksheetSubHeaders(headerRow2, "TH");
            //SetMarksheetSubHeaders(headerRow2, "PR");


            //headerRow1.Cells[3].ColSpan = 2;

            //table.Rows[0].Cells[3].ColSpan = 2;

        }

        private void SetMarksheetHeaders(Aspose.Pdf.Row row, string headerText)
        {
            Aspose.Pdf.Cell cell = row.Cells.Add();
            TextFragment text = new TextFragment(headerText);

            if (headerText == "Full Marks")
            {
                cell.ColSpan = 2;
                TextFragment text1 = new TextFragment(headerText);
                text1.TextState.FontSize = 10;
                text1.TextState.FontStyle = FontStyles.Bold;
                text1.HorizontalAlignment = HorizontalAlignment.Center;

                cell.BackgroundColor = Color.IndianRed;
                cell.Paragraphs.Add(text);
                cell.Border = new Aspose.Pdf.BorderInfo(Aspose.Pdf.BorderSide.Box, 2f, Aspose.Pdf.Color.Red);
            }
            else
            {

                text.TextState.FontSize = 10;
                text.TextState.FontStyle = FontStyles.Bold;
                text.HorizontalAlignment = HorizontalAlignment.Center;
                text.VerticalAlignment = VerticalAlignment.Top;

                cell.Paragraphs.Add(text);
                cell.Border = new Aspose.Pdf.BorderInfo(Aspose.Pdf.BorderSide.Box, 0.1f, Aspose.Pdf.Color.Black);
            }


           

        }
        private void SetMarksheetSubHeaders(Aspose.Pdf.Row row, string headerText)
        {
            Aspose.Pdf.Cell cell = row.Cells.Add();
            cell.BackgroundColor = Color.Aqua;
            row.FixedRowHeight = 20;

            TextFragment text = new TextFragment(headerText);
            text.TextState.FontSize = 10;
            text.TextState.FontStyle = FontStyles.Bold;
            text.HorizontalAlignment = HorizontalAlignment.Center;

            text.Margin.Top = 5;
            cell.Paragraphs.Add(text);
            cell.Border = new Aspose.Pdf.BorderInfo(Aspose.Pdf.BorderSide.All, 0.1f, Aspose.Pdf.Color.Black);

        }

        private void AddMarksRow(Aspose.Pdf.Table table, string studentId, string subject, string creditHour,
                                       string fullMarks, string passMarks, int obtainedMarksTh, int obtainedMarksPr
                                       )
        {

            Aspose.Pdf.Row row = table.Rows.Add();
            AddMarksDetail(row, studentId);
            AddMarksDetail(row, subject);
            AddMarksDetail(row, creditHour);
            AddMarksDetail(row, "80");
            AddMarksDetail(row, "20");

            AddMarksDetail(row, passMarks);
            if (obtainedMarksTh < 40)
            {
                AddFailedMarksDetail(row, obtainedMarksTh.ToString(), true);
            }
            else
            {
                AddMarksDetail(row, obtainedMarksTh.ToString());
            }
            AddMarksDetail(row, obtainedMarksPr.ToString());
            AddMarksDetail(row, (obtainedMarksTh + obtainedMarksPr).ToString());


        }

        private void AddFailedMarksDetail(Aspose.Pdf.Row row, string value, bool isFailed)
        {
            Aspose.Pdf.Cell cell = row.Cells.Add();

            TextFragment text = new TextFragment(value + "*");
            text.TextState.FontSize = 12;
            text.Margin.Top = 15;
            text.Margin.Bottom = 5;
            text.Margin.Left = 5;
            text.Margin.Right = 5;
            text.HorizontalAlignment = HorizontalAlignment.Center;
            text.VerticalAlignment = VerticalAlignment.Center;
            text.TextState.ForegroundColor = Aspose.Pdf.Color.Red;
            cell.Paragraphs.Add(text);

            cell.Border = new Aspose.Pdf.BorderInfo(Aspose.Pdf.BorderSide.All, 0.1f, Aspose.Pdf.Color.Black);
        }
        private void AddMarksDetail(Aspose.Pdf.Row row, string value)
        {
            Aspose.Pdf.Cell cell = row.Cells.Add();
            TextFragment text = new TextFragment(value);
            text.TextState.FontSize = 12;
            text.Margin.Top = 15;
            text.Margin.Bottom = 5;
            text.Margin.Left = 5;
            text.Margin.Right = 5;
            text.HorizontalAlignment = HorizontalAlignment.Center;
            text.VerticalAlignment = VerticalAlignment.Center;

            cell.Paragraphs.Add(text);
            cell.Border = new Aspose.Pdf.BorderInfo(Aspose.Pdf.BorderSide.All, 0.1f, Aspose.Pdf.Color.Black);

        }


        public List<List<string>> ExtractFirstRowsData(List<IFormFile> files)
        {
            List<List<string>> allFirstRowsData = new List<List<string>>();

            foreach (var file in files)
            {
                using (var stream = file.OpenReadStream())
                {
                    Workbook workbook = new Workbook(stream);

                    Worksheet worksheet = workbook.Worksheets[0];

                    Aspose.Cells.RowCollection rows = worksheet.Cells.Rows;

                    List<string> firstRowData = new List<string>();

                    foreach (Aspose.Cells.Cell cell in rows[1])
                    {
                        firstRowData.Add(cell.StringValue); 
                    }

                    allFirstRowsData.Add(firstRowData);
                }
            }

            return allFirstRowsData;
        }

        
    }
}
