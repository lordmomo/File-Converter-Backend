using FileConversion.Entity;

namespace FileConversion.Service.ServiceInterface
{
    public interface DocumentConverterInterface
    {
        ResponseMessage CreateExcelToWordAndPdfConversion(List<IFormFile> files);
        ResponseMessage CreatePdfToWordConversion(List<IFormFile> files);
        ResponseMessage CreateWordToPdfConversion(List<IFormFile> files);
        ResponseMessage mergeFiles(List<IFormFile> files);
        Task<ResponseMessage> ReplaceTextInDocuments(List<IFormFile> files, string oldText, string newText);
    }
}
