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

        private FileEditorDocx()
        {

        }

        private static FileEditorDocx _instance;

        public static FileEditorDocx GetInstance()
        {
            if (_instance == null)
            {
                _instance = new FileEditorDocx();
            }
            return _instance;
        }

        public static void EditFile(string inputFilePath, string outputFilePath)
        {

            doc = DocX.Load(inputFilePath);
            f = new Formatting();
            f.FontFamily = new Font("Arial");
            f.Size = 9D;
            tab = doc.Tables[1];
            
            FillCells();
            FormatCells();
            MergeCells();

            doc.SaveAs(outputFilePath);

        }

        private static void FillCells()
        {
            tab.Rows[1].Cells[1].Paragraphs.First().Append("Документ 1", f);
            tab.Rows[1].Cells[2].Paragraphs.First().Append("Расписка", f);
            tab.Rows[1].Cells[3].Paragraphs.First().Append("2", f);
            tab.Rows[1].Cells[4].Paragraphs.First().Append("Комментарий 1", f);

            tab.Rows[4].Cells[1].Paragraphs.First().Append("Документ 2", f);
            tab.Rows[4].Cells[2].Paragraphs.First().Append("Расписка", f);
            tab.Rows[4].Cells[3].Paragraphs.First().Append("1", f);
            tab.Rows[4].Cells[4].Paragraphs.First().Append("Комментарий 2", f);

            tab.Rows[7].Cells[1].Paragraphs.First().Append("Документ 3", f);
            tab.Rows[7].Cells[2].Paragraphs.First().Append("Расписка", f);
            tab.Rows[7].Cells[3].Paragraphs.First().Append("1", f);
            tab.Rows[7].Cells[4].Paragraphs.First().Append("Комментарий 3", f);
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
    }
}
