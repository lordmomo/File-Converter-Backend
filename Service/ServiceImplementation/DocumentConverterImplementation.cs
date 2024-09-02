using Aspose.Cells;
using Aspose.Pdf.Operators;
using Aspose.Pdf.Text;
using Aspose.Words;
using Aspose.Words.Replacing;
using FileConversion.Entity;
using FileConversion.Service.ServiceInterface;
using FileConversion.Utils;
using FileConversion.Utils.Constants;
using Microsoft.AspNetCore.Mvc;
using System.Text.RegularExpressions;

namespace FileConversion.Service.ServiceImplementation
{

    public class DocumentConverterImplementation : DocumentConverterInterface
    {
        public ResponseMessage CreateWordToPdfConversion(List<IFormFile> files)
        {

            var generatedPdfPaths = new List<string>();
            if (files == null || files.Count == 0)
            {
                return new ResponseMessage { Success = false, Message = Message.fileNotProvided };
            }

            foreach (var file in files)
            {
                if (file.Length > 0)
                {
                    var filePath = Path.GetTempFileName();

                    var realFilename = file.FileName;
                    using (var stream = new FileStream(filePath, FileMode.Create))
                    {
                        file.CopyTo(stream);
                    }

                    if (file.FileName.EndsWith(".docx"))
                    {
                        ConvertDocxToPdf(filePath, Constants.wordToPdfStoragePath, realFilename);
                        string outputPdfPath = Path.Combine(Constants.wordToPdfStoragePath, $"{Path.GetFileNameWithoutExtension(realFilename)}.pdf");
                        generatedPdfPaths.Add(outputPdfPath);

                    }
                    else
                    {
                        return new ResponseMessage { Success = false, Message = Message.onlyWordFilesAllowed };

                    }


                    System.IO.File.Delete(filePath);
                }
            }

            return new ResponseMessage { Success = true, Message = Message.successfullyTransformedWordToPdf, Data = generatedPdfPaths };

        }

        public ResponseMessage CreatePdfToWordConversion(List<IFormFile> files)
        {

            var messages = new List<string>();

            var generatedPdfPaths = new List<string>();
            if (files == null || files.Count == 0)
            {
                return new ResponseMessage { Success = false, Message = Message.fileNotProvided };
            }

            foreach (var file in files)
            {
                if (file.Length > 0)
                {
                    var filePath = Path.GetTempFileName();

                    var realFilename = file.FileName;
                    using (var stream = new FileStream(filePath, FileMode.Create))
                    {
                        file.CopyTo(stream);
                    }

                    if (file.FileName.EndsWith(".pdf"))
                    {
                        ConvertPdfToDocx(filePath, Constants.pdfToWordStoragePath, realFilename);
                        string outputPdfPath = Path.Combine(Constants.pdfToWordStoragePath, $"{Path.GetFileNameWithoutExtension(realFilename)}.docx");
                        generatedPdfPaths.Add(outputPdfPath);
                    }
                    else
                    {
                        messages.Add($"File '{file.FileName}' could not be converted. Only Pdf files are supported.");
                        return new ResponseMessage { Success = false, Message = Message.onlyPdfFilesAllowed };

                    }


                    System.IO.File.Delete(filePath);
                }
            }

            return new ResponseMessage { Success = true, Message = Message.successfullyTransformedPdfToWord, Data = generatedPdfPaths };

        }

        public ResponseMessage CreateExcelToWordAndPdfConversion(List<IFormFile> files)
        {
            var messages = new List<string>();
            var generatedPdfPaths = new List<string>();

            if (files == null || files.Count == 0)
            {
                return new ResponseMessage { Success = false, Message = Message.fileNotProvided };
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

                    if (file.FileName.EndsWith(".xlsx"))
                    {
                        ConvertExcelToDocxAndPdf(filePath, Constants.excelToWordAndPdfStoragePath, file.FileName);

                        string outputPdfPath = Path.Combine(Constants.excelToWordAndPdfStoragePath, $"{Path.GetFileNameWithoutExtension(file.FileName)}.pdf");
                        generatedPdfPaths.Add(outputPdfPath);
                        string outputWordPath = Path.Combine(Constants.excelToWordAndPdfStoragePath, $"{Path.GetFileNameWithoutExtension(file.FileName)}.docx");
                        generatedPdfPaths.Add(outputWordPath);

                    }
                    else
                    {
                        return new ResponseMessage { Success = false, Message = Message.onlyExcelFilesAllowed };

                    }


                    System.IO.File.Delete(filePath);
                }
            }

            return new ResponseMessage { Success = true, Message = Message.successfulltTransformedExcelToWordAndPdf, Data = generatedPdfPaths };

        }

        public ResponseMessage mergeFiles(List<IFormFile> files)
        {
            try
            {


                if (files == null || files.Count == 0)
                    return new ResponseMessage { Success = false, Message = Message.fileNotProvided };

                List<string> outputFilePath = new List<string>();

                Aspose.Pdf.Document document = new Aspose.Pdf.Document();
                foreach (var file in files)
                {
                    if (file.Length > 0)
                    {
                        var filePath = Path.GetTempFileName();

                        using (var stream = new FileStream(filePath, FileMode.Create))
                        {
                            file.CopyTo(stream);
                        }
                        if (!file.FileName.EndsWith(".pdf"))
                        {
                            return new ResponseMessage { Success = false, Message = Message.onlyPdfFilesAllowed };

                        }

                        Aspose.Pdf.Document doc = new Aspose.Pdf.Document(filePath);

                        if (file.FileName.EndsWith(".pdf"))
                        {
                            foreach (Aspose.Pdf.Page page in doc.Pages)
                            {
                                document.Pages.Add(page);
                            }

                        }
                        else
                        {
                            return new ResponseMessage { Success = false, Message = Message.onlyPdfFilesAllowed };

                        }


                        System.IO.File.Delete(filePath);
                    }
                }

                var outputFile = Path.Combine(Constants.mergedPdfsStoragePath, "merged_file.pdf");
                outputFilePath.Add(outputFile);
                document.Save(outputFile);

                return new ResponseMessage { Success = true, Message = Message.successfullyMergedProvidedPdfs, Data = outputFilePath };
            }
            catch (IndexOutOfRangeException ex)
            {
                Console.WriteLine($"IndexOutOfRangeException in mergeFiles method: {ex.Message}");
                return new ResponseMessage { Success = false, Message = Message.unexpectedErrorWhileMergingFiles };
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Exception in mergeFiles method: {ex.Message}");

                return new ResponseMessage { Success = false, Message = Message.unexpectedErrorWhileProcessingFiles };

            }
        }



        private void ConvertExcelToDocxAndPdf(string inputFilePath, string outputDirectory, string originalFileName)
        {
            Workbook workbook = new Workbook(inputFilePath);

            string outputDocxPath = Path.Combine(outputDirectory + Path.GetFileNameWithoutExtension(originalFileName) + ".docx");
            string outputPdfPath = Path.Combine(outputDirectory + Path.GetFileNameWithoutExtension(originalFileName) + ".pdf");

            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                builder.Write(worksheet.Name);
                builder.Writeln();

                builder.StartTable();
                foreach (Row row in worksheet.Cells.Rows)
                {
                    foreach (Cell cell in row)
                    {
                        builder.InsertCell();
                        builder.Write(cell.StringValue);
                    }
                    builder.EndRow();

                }
                builder.EndTable();
            }

            doc.Save(outputDocxPath);

            workbook.Save(outputPdfPath, Aspose.Cells.SaveFormat.Pdf);
        }

        private void ConvertPdfToDocx(string inputFilePath, string outputDirectory, string realFilename)
        {
            string outputFilePath = Path.Combine(outputDirectory, Path.GetFileNameWithoutExtension(realFilename) + ".docx");

            Document doc = new Document(inputFilePath);
            doc.Save(outputFilePath, Aspose.Words.SaveFormat.Docx);

        }

        private void ConvertDocxToPdf(string inputFilePath, string outputDirectory, string originalFilename)
        {
            Document doc = new Document(inputFilePath);
            string outputFilePath = Path.Combine(outputDirectory, Path.GetFileNameWithoutExtension(originalFilename) + ".pdf");
            doc.Save(outputFilePath, Aspose.Words.SaveFormat.Pdf);
        }

        public async Task<ResponseMessage> ReplaceTextInDocuments(List<IFormFile> files, string oldText, string newText)
        {

            List<string> outputFilePathList = new List<string>();


            if (files == null || files.Count == 0)
            {
                return new ResponseMessage { Success = false, Message = Message.fileNotProvided };
            }

            if (string.IsNullOrEmpty(oldText) || string.IsNullOrEmpty(newText))
            {
                return new ResponseMessage { Success = false, Message = Message.textReplacementWordsNotProvided };

            }

            try
            {
                bool anyReplacementsMade = false;

                foreach (var file in files)
                {
                    if (file.Length == 0)
                    {
                        continue;
                    }

                    string fileExtension = Path.GetExtension(file.FileName).ToLower();
                    string tempFilePath = Path.GetTempFileName();
                    string outputFilePath = Path.Combine(Constants.replacedTextStoragePath, Path.GetFileNameWithoutExtension(file.FileName) + "_Modified" + fileExtension);

                    using (var fileStream = file.OpenReadStream())
                    {
                        using (var tempFileStream = new FileStream(tempFilePath, FileMode.Create))
                        {
                            fileStream.CopyTo(tempFileStream);
                        }

                    }
                    bool fileProcessed = false;

                    try
                    {



                        if (fileExtension == ".docx" || fileExtension == ".doc")
                        {
                            fileProcessed = await Task.Run(() => ReplaceTextInWordDocument(tempFilePath, oldText, newText, outputFilePath));
                        }
                        else if (fileExtension == ".xlsx" || fileExtension == ".xls")
                        {
                            fileProcessed = await Task.Run(() => ReplaceTextInExcelDocument(tempFilePath, oldText, newText, outputFilePath));
                        }
                        else if (fileExtension == ".pdf")
                        {
                            fileProcessed = await Task.Run(() => ReplaceTextInPdfDocument(tempFilePath, oldText, newText, outputFilePath));
                        }
                        else
                        {
                            return new ResponseMessage { Success = false, Message = Message.unsupportedFileType };

                        }

                    }
                    catch (Exception ex)
                    {

                    }
                    anyReplacementsMade = fileProcessed;
                    outputFilePathList.Add(outputFilePath);
                    //return new ResponseMessage { Success = true, Message = Message.successfullyReplacedTexts, Data = outputFilePathList };
                }
                if (anyReplacementsMade)
                {
                    return new ResponseMessage { Success = true, Message = Message.successfullyReplacedTexts, Data = outputFilePathList };

                }
                else
                {
                    return new ResponseMessage { Success = false, Message = Message.textToReplaceNotFound };

                }

            }
            catch (Exception ex)
            {

                Console.WriteLine(ex.Message);
                return new ResponseMessage { Success = false, Message = Message.unexpectedErrorWhileProcessingFiles };

            }

        }


        private bool ReplaceTextInWordDocument(string inputFilePath, string oldText, string newText, string outputFilePath)
        {
            bool textReplaced = false;

            Document doc = new Document(inputFilePath);
            FindReplaceOptions options = new FindReplaceOptions
            {
                MatchCase = true,
                FindWholeWordsOnly = true,
                Direction = FindReplaceDirection.Forward
            };
            textReplaced = doc.Range.Replace(oldText, newText, options) > 0;

            doc.Save(outputFilePath);
            return textReplaced;


        }

        private bool ReplaceTextInExcelDocument(string inputFilePath, string oldText, string newText, string outputFilePath)
        {
            bool textReplaced = false;


            Workbook workbook = new Workbook(inputFilePath);
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                Cells cells = worksheet.Cells;
                foreach (Cell cell in cells)
                {
                    if (cell.Value != null && cell.Value.ToString().Contains(oldText))
                    {
                        cell.Value = cell.Value.ToString().Replace(oldText, newText);
                        textReplaced = true;
                    }
                }
            }
            workbook.Save(outputFilePath);
            return textReplaced;

        }

        private bool ReplaceTextInPdfDocument(string inputFilePath, string oldText, string newText, string outputFilePath)
        {
            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(inputFilePath);

            bool textReplaced = false;
            TextFragmentAbsorber textFragmentAbsorber = new TextFragmentAbsorber();
            foreach (Aspose.Pdf.Page page in pdfDocument.Pages)
            {
                page.Accept(textFragmentAbsorber);

                TextFragmentCollection textFragments = textFragmentAbsorber.TextFragments;

                string pattern = $@"\b{Regex.Escape(oldText)}\b";
                Regex regex = new Regex(pattern, RegexOptions.IgnoreCase);

                foreach (TextFragment textFragment in textFragments)
                {
                    if (regex.IsMatch(textFragment.Text))
                    {
                        textFragment.Text = regex.Replace(textFragment.Text, newText);
                        textReplaced = true;
                    }
                }
            }

            pdfDocument.Save(outputFilePath);
            return textReplaced;
            
        }

    }
}
