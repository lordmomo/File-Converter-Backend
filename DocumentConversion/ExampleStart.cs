using Aspose.Pdf;
using Aspose.Pdf.Text;
using FileConversion.Entity;

namespace FileConversion.DocumentConversion
{
    public static class ExampleStart
    {
        private static readonly string _dataDir = "C:\\Users\\i44375\\source\\repos\\FileConversion\\Utils\\Samples";

        public static void HelloWorld()
        {
            Document document = new Document();
            Page page = document.Pages.Add();

            page.Paragraphs.Add(new TextFragment("Hello world"));

            var outputFileName = Path.Combine(_dataDir, "HelloWorld_out.pdf");
            document.Save(outputFileName);
        }

        public static void Greeting(Pesron person)
        {
            Document document = new Document();
            Page page = document.Pages.Add();

            //page.Paragraphs.Add(new Aspose.Pdf.Text.TextFragment("Hello world"));

            TextFragment textFragment = new TextFragment();
            textFragment.Text = $"Hello {person.Name},\n\n";
            textFragment.Text += $"Address: {person.Address}\n";
            textFragment.Text += $"Age: {person.Age}\n\n";
            textFragment.Text += "Welcome to our service!";

            // Set formatting options for the text fragment
            textFragment.TextState.FontSize = 12;
            textFragment.TextState.FontStyle = FontStyles.Bold;

            // Add the text fragment to the page
            page.Paragraphs.Add(textFragment);

            var outputFileName = Path.Combine(_dataDir, "GreetingModified_out.pdf");
            document.Save(outputFileName);
        }

    }
}
