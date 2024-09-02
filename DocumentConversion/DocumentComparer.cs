using Aspose.Pdf.Text;
using Aspose.Pdf;
using FileConversion.Entity;

namespace FileConversion.DocumentConversion
{
    public  static class DocumentComparer
    {
        public static string CompareDocument(CompareDocumentsDto documentsDto)
        {

            string filePath1 = @"C:\Users\i44375\source\repos\FileConversion\Utils\Samples\" + documentsDto.Document1 + ".pdf";
            string filePath2 = @"C:\Users\i44375\source\repos\FileConversion\Utils\Samples\" + documentsDto.Document2 + ".pdf";

            string outputFile = @"C:\Users\i44375\source\repos\FileConversion\Utils\Samples\" + documentsDto.OutFile + ".pdf";

            Document document1 = new Document(filePath1);
            Document document2 = new Document(filePath2);


            TextAbsorber textAbsorber1 = new TextAbsorber();
            TextAbsorber textAbsorber2 = new TextAbsorber();

            document1.Pages.Accept(textAbsorber1);
            string text1 = textAbsorber1.Text;

            document2.Pages.Accept(textAbsorber2);
            string text2 = textAbsorber2.Text;

            if(text1.Equals(text2))
            {
                Console.WriteLine("The Pdfs are same");
               
                return "file are identical";
            }
            else
            {

                TextFragmentAbsorber textFragmentAbsorber1 = new TextFragmentAbsorber();
                TextFragmentAbsorber textFragmentAbsorber2 = new TextFragmentAbsorber();
                textFragmentAbsorber1.TextSearchOptions = new TextSearchOptions(true); // Enable text comparison
                textFragmentAbsorber2.TextSearchOptions = new TextSearchOptions(true); // Enable text comparison

                // Visit the documents with the text fragment absorbers
                document1.Pages.Accept(textFragmentAbsorber1);
                document2.Pages.Accept(textFragmentAbsorber2);

                // Iterate through text fragments in doc2 and compare with doc1
                foreach (TextFragment fragment2 in textFragmentAbsorber2.TextFragments)
                {
                    // Find corresponding fragment in doc1
                    TextFragment correspondingFragment1 = FindCorrespondingFragment(fragment2, textFragmentAbsorber1);

                    if (correspondingFragment1 != null)
                    {
                        // Compare text content
                        if (!fragment2.Text.Equals(correspondingFragment1.Text))
                        {
                            // Highlight the text fragment in doc2 with color
                            fragment2.TextState.BackgroundColor = Color.Yellow; // Set the background color
                            fragment2.TextState.ForegroundColor = Color.Red; // Set the foreground (text) color
                        }
                    }
                }

                // Save the modified document with highlighted differences
                document2.Save(outputFile);
                return "success change file saved";

            }

        }

        static TextFragment FindCorrespondingFragment(TextFragment fragment, TextFragmentAbsorber textFragmentAbsorber)
        {
            // Initialize TextFragmentAbsorber to find the corresponding fragment in the collection
            textFragmentAbsorber.TextSearchOptions = new TextSearchOptions(true); // Enable text comparison
            textFragmentAbsorber.Visit(fragment.Page);

            // Return the first fragment found (assuming there's only one match)
            if (textFragmentAbsorber.TextFragments.Count > 0)
            {
                return textFragmentAbsorber.TextFragments[1]; // Ensure index is within range (1-based index)
            }

            return null;
        }

        static TextFragment GetCorrespondingFragment(TextFragment fragment, Document document)
        {
            // Initialize TextFragmentAbsorber to find the corresponding fragment in the document
            TextFragmentAbsorber textFragmentAbsorber = new TextFragmentAbsorber(fragment.Text);
            textFragmentAbsorber.TextSearchOptions = new TextSearchOptions(true); // Enable text comparison
            textFragmentAbsorber.Visit(document);

            // Return the first fragment found (assuming there's only one match)
            if (textFragmentAbsorber.TextFragments.Count > 0)
            {
                return textFragmentAbsorber.TextFragments[1];
            }

            return null;
        }
    }
}
