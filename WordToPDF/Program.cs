using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocToPDFConverter;
using Syncfusion.Pdf;

namespace WordToPDF
{
    class Program
    {
        static void Main(string[] args)
        {
            //Get the path of existing Word document
            string fullpath = @"...\..\DocToPDF.docx";

            //Loads an existing Word document
            WordDocument wordDocument = new WordDocument(fullpath, FormatType.Docx);
            
            //Creates an instance of the DocToPDFConverter
            DocToPDFConverter converter = new DocToPDFConverter();

            //Converts Word document into PDF document
            PdfDocument pdfDocument = converter.ConvertToPDF(wordDocument);
            
            //Releases all resources used by DocToPDFConverter
            converter.Dispose();

            //Saves the PDF file 
            pdfDocument.Save("DocToPDF.pdf");

            //Closes the instance of document objects
            pdfDocument.Close(true);
            wordDocument.Close();
        }
    }
}
