using Aspose.Cells;
using Aspose.Pdf.Text;
using Aspose.Words;
using FileConversion.Entity;
using FileConversion.Service.ServiceInterface;
using FileConversion.Utils;
using Microsoft.AspNetCore.Mvc;
using System.IO.Compression;

namespace FileConversion.Controllers
{
    public class DocumentConverterController : Controller
    {
        private readonly DocumentConverterInterface _documentConverterInterface;
        public DocumentConverterController(DocumentConverterInterface documentConverterInterface) {
            _documentConverterInterface = documentConverterInterface;
        }
        public IActionResult Index()
        {
            return View();
        }

        [HttpPost(Urls.convertWordToPdfApi)]
        public IActionResult UploadWordFiles(List<IFormFile> files)
        {
            ResponseMessage response = _documentConverterInterface.CreateWordToPdfConversion(files);
            return Json(response);
        }

        [HttpPost(Urls.convertPdfToWordApi)]
        public IActionResult UploadPdfFiles(List<IFormFile> files)
        {

            ResponseMessage response = _documentConverterInterface.CreatePdfToWordConversion(files);
            return Json(response);

        }

        [HttpPost(Urls.convertExcelToWordAndExcelApi)]
        public IActionResult UploadExcel(List<IFormFile> files)
        {
            ResponseMessage response = _documentConverterInterface.CreateExcelToWordAndPdfConversion(files);
            return Json(response);
        }


        [HttpGet(Urls.downloadPdfsInZipApi)]
        public IActionResult DownloadConvertedExcelZip([FromQuery] List<string> pdfPaths)
        {
            try
            {
                using (var memoryStream = new MemoryStream())
                {
                    using (var archive = new ZipArchive(memoryStream, ZipArchiveMode.Create, true))
                    {
                        foreach (var pdfPath in pdfPaths)
                        {
                            var fileInfo = new FileInfo(pdfPath);
                            var entry = archive.CreateEntry(fileInfo.Name);

                            using (var fileStream = new FileStream(pdfPath, FileMode.Open, FileAccess.Read))
                            using (var entryStream = entry.Open())
                            {
                                fileStream.CopyTo(entryStream);
                            }
                        }
                    }

                    return File(memoryStream.ToArray(), "application/zip", "ConvertedExcelFile.zip");
                }
            }
            catch (Exception ex)
            {
                return StatusCode(500, new ResponseMessage { Success = false, Message = $"Failed to download file: {ex.Message}" });
            }
        }



        [HttpGet(Urls.downloadSingleFile)]
        public IActionResult DownloadSingleFile([FromQuery] string filePath)
        {

            if (!System.IO.File.Exists(filePath))
            {
                return NotFound();
            }

            try
            {
                string contentType;
                if (Path.GetExtension(filePath).Equals(".pdf", StringComparison.InvariantCultureIgnoreCase))
                {
                    contentType = "application/pdf";
                }
                else if (Path.GetExtension(filePath).Equals(".docx", StringComparison.InvariantCultureIgnoreCase))
                {
                    contentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
                }
                else
                {
                    // Unsupported file type
                    return BadRequest(new { message = "Unsupported file type." });
                }

                var fileBytes = System.IO.File.ReadAllBytes(filePath);
                return File(fileBytes, contentType, Path.GetFileName(filePath));
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { message = $"Failed to download file: {ex.Message}" });
            }
        }

        [HttpPost(Urls.mergePdfsApi)]
        public IActionResult MergeFiles( List<IFormFile> files)
        {
            ResponseMessage response = _documentConverterInterface.mergeFiles(files);
            return Json(response);
        }

        [HttpGet(Urls.mergePdfsApi)]
        public IActionResult DownloadMergedPdf([FromQuery] string filePath)
        {
            if (!System.IO.File.Exists(filePath))
            {
                return NotFound();
            }

            try
            {
                string contentType;
                if (Path.GetExtension(filePath).Equals(".pdf", StringComparison.InvariantCultureIgnoreCase))
                {
                    contentType = "application/pdf";
                }
                else
                {
                    // Unsupported file type
                    return BadRequest(new { message = "Unsupported file type." });
                }

                var fileBytes = System.IO.File.ReadAllBytes(filePath);
                return File(fileBytes, contentType, Path.GetFileName(filePath));
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { message = $"Failed to download file: {ex.Message}" });
            }
        }



        [HttpPost("replaceText")]
        public async Task<IActionResult> ReplaceTextInDocuments([FromForm] List<IFormFile> files, [FromForm] string oldText, [FromForm] string newText)
        {
          

            var response = await _documentConverterInterface.ReplaceTextInDocuments(files, oldText, newText);
            return Json(response);
        }

    }
}

