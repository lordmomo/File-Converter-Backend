namespace FileConversion.Utils.Constants
{
    public static class Message
    {
        public const string fileNotProvided = "Files are not provided";
        public const string onlyExcelFilesAllowed = "Only Excel files are supported";
        public const string onlyPdfFilesAllowed = "Only Pdf files are supported";
        public const string onlyWordFilesAllowed = "Only Word Docx files are supported";

        public const string textReplacementWordsNotProvided = "Text to replace and new text must be provided.";

        public const string successfullyReplacedTexts = "Text replacement successful";
        public const string successfullySavePdfs = "Successfully saved PDFs";

        public const string successfullyTransformedWordToPdf = "Successfully transformed the provided word file";
        public const string successfullyTransformedPdfToWord = "Successfully transformed the provided pdf file";
        public const string successfulltTransformedExcelToWordAndPdf = "Successfully transformed the provided excel file";
        public const string successfullyMergedProvidedPdfs = "Successfully merged the provided pdf files";
        public const string mergePdfError = " Error merging pdf files";

        public const string unsupportedFileType = "Unsupported file type.";
        public const string unexpectedErrorWhileMergingFiles ="An unexpected error occurred while merging your files.";
        public const string unexpectedErrorWhileProcessingFiles ="An unexpected error occurred while processing your files.";

        public const string unmatchedInputFileForTemplate = "Unmatched input file caused an error when creating the template.";
        public const string textToReplaceNotFound = "Text to replace is not present in file";
    }
}
