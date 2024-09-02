using Aspose.Cells;
using Aspose.Pdf;
using Aspose.Pdf.Text;
using Aspose.Words.Pdf2Word.FixedFormats;
using Aspose.Words.Tables;
using FileConversion.Entity;
using FileConversion.Service.ServiceInterface;
using FileConversion.Utils;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using System.IO.Compression;
using System.Reflection.Metadata;

namespace FileConversion.Controllers;

public class ReportGeneratorController : Controller
{

    private readonly ReportGeneratorInterface _reportGeneratorInterface;

    public ReportGeneratorController(ReportGeneratorInterface reportGeneratorInterface)
    {
        _reportGeneratorInterface = reportGeneratorInterface;
    }

    public IActionResult Index()
    {
        return View();
    }

    [HttpPost]
    [Route(Urls.generateInvoiceApi)]
    public IActionResult GenerateReport([FromForm] List<IFormFile> files)
    {

        ReportGenerationResponseMessage response = _reportGeneratorInterface.CreateReport(files);
    
        return Json(response);

    }

    //[HttpGet(Urls.downloadInvoicesInZipApi)]
    //public IActionResult DownloadInvoicesZip([FromQuery] List<string> pdfPaths)
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
    //            return File(memoryStream.ToArray(), "application/zip", "GeneratedFiles.zip");
    //        }
    //    }
    //    catch (Exception ex)
    //    {
    //        // Handle exceptions and return appropriate response
    //        return StatusCode(500, new ResponseMessage { Success = false, Message = $"Failed to download PDFs: {ex.Message}" });
    //    }
    //}


    //[HttpGet(Urls.downloadSingleInvoiceInPdf)]
    //public IActionResult DownloadSinglePdf([FromQuery] string pdfPath)
    //{
    //    var filePath = pdfPath;

    //    if (!System.IO.File.Exists(filePath))
    //    {
    //        return NotFound();
    //    }

    //    try
    //    {
    //        var fileBytes = System.IO.File.ReadAllBytes(filePath);
    //        return File(fileBytes, "application/pdf", Path.GetFileName(filePath));
    //    }
    //    catch (Exception ex)
    //    {
    //        return StatusCode(500, new { message = $"Failed to download PDF: {ex.Message}" });
    //    }
    //}



}
