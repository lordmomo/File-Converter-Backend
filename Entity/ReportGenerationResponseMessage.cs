namespace FileConversion.Entity
{
    public class ReportGenerationResponseMessage
    {
        public Boolean Success { get; set; }
        public string Message { get; set; }

        public List<string> Data { get; set; }
        public List<string> PreviewData { get; set; }

    }
}
