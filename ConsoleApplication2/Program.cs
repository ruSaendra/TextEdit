using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace ConsoleApplication2
{
    class Program
    {
        static void Main(string[] args)
        {
            // string inputFilename = @"C:\Users\Saendra\Google Drive\ConsultationResult.docx";
            // string outputFilename = @"C:\Users\Saendra\Google Drive\ConsultationResultEdited.docx";
            string inputFilename = @"ConsultationResult.docx";
            string outputFilenameDocX = @"ConsultationResultEditedDocX.docx";
            string outputFilenameXml = @"ConsultationResultEditedXml.docx";
                        
            if(File.Exists(inputFilename))
            {
                FileEditorDocx.EditFile(inputFilename, outputFilenameDocX);
                FileEditorOpenXml.EditFile(inputFilename, outputFilenameXml);
                
            }            
        }
    }
}
