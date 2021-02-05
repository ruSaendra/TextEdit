using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace ConsoleApplication2
{
    class FileEditorDocx
    {
        private static DocX doc;
        private static Formatting f;
        private static Table tab;

        public static void EditFile(string inputFilePath, string outputFilePath)
        {

            if(LoadFile(inputFilePath))
            {
                f = new Formatting();
                f.FontFamily = new Font("Arial");
                f.Size = 9D;
                tab = doc.Tables[1];

                FillCells();
                FormatCells();
                MergeCells();

                try
                {
                    doc.SaveAs(outputFilePath);
                }
                catch (Exception e)
                {
                    Console.WriteLine(e);
                }
            }           

        }

        private static Boolean LoadFile(string inputFilePath)                           // Attempts to load the file
        {
            try
            {
                doc = DocX.Load(inputFilePath);
            }
            catch (Exception e)
            {
                return false;
            }
            return true;
        }

        private static void FillCells()                                                 // Fills required cells with text
        {
            FillRow(1, "Документ 1", "Расписка", 2, "Комментарий 1");
            FillRow(4, "Документ", "Расписка", 1, "Комментарий 2");
            FillRow(7, "Документ", "Расписка", 1, "Комментарий 3");
        }

        private static void FormatCells()                                               // Changes formatting of cells
        {
            tab.Rows[1].Cells[1].Paragraphs.First().Bold().Italic();
            tab.Rows[1].Cells[2].Paragraphs.First().Italic();
            tab.Rows[4].Cells[1].Paragraphs.First().StrikeThrough(StrikeThrough.strike);
            tab.Rows[4].Cells[4].Paragraphs.First().Bold();

        }

        private static void MergeCells()                                                // Merges specific cells
        {
            tab.MergeCellsInColumn(0, 1, 3);
            tab.MergeCellsInColumn(0, 4, 6);
        }

        private static void FillRow(int row, string docName, string docType, int qty, string comment)        // Inputs text into cells in a row
        {
            tab.Rows[row].Cells[1].Paragraphs.First().Append(docName, f);
            tab.Rows[row].Cells[2].Paragraphs.First().Append(docType, f);
            tab.Rows[row].Cells[3].Paragraphs.First().Append(qty.ToString(), f);
            tab.Rows[row].Cells[4].Paragraphs.First().Append(comment, f);
        }
    }
}
