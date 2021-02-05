using System.IO;

namespace ConsoleApplication2
{
    class Program
    {
        static void Main(string[] args)
        {
            string inputFilename = @"ConsultationResult.docx";
            string outputFilenameDocX = @"ConsultationResultEditedDocX.docx";
            string outputFilenameXml = @"ConsultationResultEditedXml.docx";
                        
            if(File.Exists(inputFilename))
            {
                FileEditorDocx.EditFile(inputFilename, outputFilenameDocX);         // Edit using Xceed DocX library
                FileEditorOpenXml.EditFile(inputFilename, outputFilenameXml);       // Edit using openXml library
                
            }            
        }
    }
}
