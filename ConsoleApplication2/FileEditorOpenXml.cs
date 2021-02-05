using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ConsoleApplication2
{
    class FileEditorOpenXml
    {
        private static WordprocessingDocument wpDoc;
        private static Body body;
        private static Table tab;
        private static RunProperties rProp;

        enum FStyle
        {
            Bold,
            Italic,
            Strike
        }

        
        public static void EditFile(string inputFilePath, string outputFilePath)
        {
            if(LoadFile(inputFilePath, outputFilePath))
            {
                wpDoc = WordprocessingDocument.Open(outputFilePath, true);

                body = wpDoc.MainDocumentPart.Document.Body;

                tab = body.Elements<Table>().ElementAt(1);

                rProp = SetFormatting("Arial", 18);

                FillCells();
                FormatCells();
                MergeCells();

                try
                {
                    wpDoc.Close();
                }
                catch (Exception e)
                {
                    Console.WriteLine(e);
                }
            }
        }

        private static Boolean LoadFile(string inputFilePath, string outputFilePath)
        {
            try
            {
                if (File.Exists(outputFilePath))
                {
                    File.Delete(outputFilePath);
                }
                File.Copy(inputFilePath, outputFilePath);
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
            InputTextIntoCell(4, 4, "Комментарий 3");

            InputTextIntoCell(7, 1, "Документ 3");
            InputTextIntoCell(7, 2, "Расписка");
            InputTextIntoCell(7, 3, "1");
            InputTextIntoCell(7, 4, "Комментарий 3");
        }

        private static void FormatCells()
        {
            FormatCell(1, 1, new FStyle[] { FStyle.Bold, FStyle.Italic } );
            FormatCell(1, 2, new FStyle[] { FStyle.Italic } );
            FormatCell(4, 1, new FStyle[] { FStyle.Strike } );
            FormatCell(4, 4, new FStyle[] { FStyle.Bold } );
        }

        private static void MergeCells()
        {
            CellsToMerge(0, 1, 3);
            CellsToMerge(0, 4, 6);
        }

        private static void InputTextIntoCell (int row, int col, string textInput)
        {
            TableRow tRow = tab.Elements<TableRow>().ElementAt(row);

            TableCell tCell = tRow.Elements<TableCell>().ElementAt(col);

            Paragraph par = tCell.Elements<Paragraph>().Count() == 0 ?
                tCell.AppendChild(new Paragraph()) :
                tCell.Elements<Paragraph>().First();

            Run run = par.Elements<Run>().Count() == 0 ?
                par.AppendChild(new Run()) :
                par.Elements<Run>().First();

            run.AppendChild(rProp.CloneNode(true));
            Text txt = run.AppendChild(new Text());
            txt.Text = textInput;
        }

        private static void FormatCell(int row, int col, FStyle[] styles)
        {
            TableRow tRow = tab.Elements<TableRow>().ElementAt(row);

            TableCell tCell = tRow.Elements<TableCell>().ElementAt(col);

            Paragraph par = tCell.Elements<Paragraph>().Count() == 0 ?
                tCell.AppendChild(new Paragraph()) :
                tCell.Elements<Paragraph>().First();

            Run run = par.Elements<Run>().Count() == 0 ?
                par.AppendChild(new Run()) :
                par.Elements<Run>().First();

            foreach (FStyle st in styles)
            {
                switch (st)
                {
                    case FStyle.Bold:
                        run.Elements<RunProperties>().First().Append(new Bold());
                        break;
                    case FStyle.Italic:
                        run.Elements<RunProperties>().First().Append(new Italic());
                        break;
                    case FStyle.Strike:
                        run.Elements<RunProperties>().First().Append(new Strike());
                        break;
                }
            }
        }

        private static void CellsToMerge(int col, int rowStart, int rowEnd)
        {
            TableCellProperties tcPropStart = new TableCellProperties();
            tcPropStart.Append(new VerticalMerge()
            {
                Val = MergedCellValues.Restart
            });

            TableCellProperties tcPropNext = new TableCellProperties();
            tcPropNext.Append(new VerticalMerge()
            {
                Val = MergedCellValues.Continue
            });

            TableCell tCell = tab.Elements<TableRow>().ElementAt(rowStart).Elements<TableCell>().ElementAt(col);
            tCell.Append(tcPropStart);

            for (int i = rowStart + 1; i <= rowEnd; i++)
            {
                tCell = tab.Elements<TableRow>().ElementAt(i).Elements<TableCell>().ElementAt(col);
                tCell.Append(tcPropNext.CloneNode(true));
            }

        }

        private static RunProperties SetFormatting(string fontStyle, int fontSize)
        {
            RunProperties rPr = new RunProperties();

            rPr.RunFonts = new RunFonts()
            {
                Ascii = fontStyle
            };

            rPr.FontSize = new FontSize()
            {
                Val = fontSize.ToString()
            };

            return rPr;
        }
    }
}
