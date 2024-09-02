
using FileConversion.Entity;

namespace FileConversion.Service.ServiceInterface
{
    public interface ReportGeneratorInterface
    {
        public ReportGenerationResponseMessage CreateReport(List<IFormFile> files);
    }
}
