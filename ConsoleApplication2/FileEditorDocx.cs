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

        private static Boolean LoadFile(string inputFilePath)
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

        private static void FillCells()
        {
            InputTextIntoCell(1, 1, "Документ 1");
            InputTextIntoCell(1, 2, "Расписка");
            InputTextIntoCell(1, 3, "2");
            InputTextIntoCell(1, 4, "Комментарий 1");

            InputTextIntoCell(4, 1, "Документ 2");
            InputTextIntoCell(4, 2, "Расписка");
            InputTextIntoCell(4, 3, "1");
            InputTextIntoCell(4, 4, "Комментарий 2");

            InputTextIntoCell(7, 1, "Документ 3");
            InputTextIntoCell(7, 2, "Расписка");
            InputTextIntoCell(7, 3, "1");
            InputTextIntoCell(7, 4, "Комментарий 3");
        }

        private static void FormatCells()
        {
            tab.Rows[1].Cells[1].Paragraphs.First().Bold().Italic();
            tab.Rows[1].Cells[2].Paragraphs.First().Italic();
            tab.Rows[4].Cells[1].Paragraphs.First().StrikeThrough(StrikeThrough.strike);
            tab.Rows[4].Cells[4].Paragraphs.First().Bold();

        }

        private static void MergeCells()
        {
            tab.MergeCellsInColumn(0, 1, 3);
            tab.MergeCellsInColumn(0, 4, 6);
        }

        private static void InputTextIntoCell(int row, int col, string textInput)
        {
            tab.Rows[row].Cells[col].Paragraphs.First().Append(textInput, f);
        }
    }
}
