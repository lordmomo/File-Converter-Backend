using FileConversion.Entity;

namespace FileConversion.Service.ServiceInterface
{
    public interface StudentReportInterfae
    {
        public ReportGenerationResponseMessage CreateStudentReport(List<IFormFile> files);
        public List<List<string>> ExtractFirstRowsData(List<IFormFile> files);
        public List<string> getPreviewPath();
    }
}
