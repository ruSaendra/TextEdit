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
                MergeCellsInColumn(0);
                // MergeCells();

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

        private static Boolean LoadFile(string inputFilePath, string outputFilePath)            // Attempts to load a file and create a copy to edit
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

        private static void FillCells()                                                         // Fills required cells with text
        {
            InputTextIntoRow(1, "Документ 1", "Расписка", 2, "Комментарий 1");
            InputTextIntoRow(4, "Документ 2", "Расписка", 1, "Комментарий 2");
            InputTextIntoRow(7, "Документ 3", "Расписка", 1, "Комментарий 3");

        }

        private static void FormatCells()                                                       // Changes formatting of cells
        {
            FormatCell(1, 1, new FStyle[] { FStyle.Bold, FStyle.Italic } );
            FormatCell(1, 2, new FStyle[] { FStyle.Italic } );
            FormatCell(4, 1, new FStyle[] { FStyle.Strike } );
            FormatCell(4, 4, new FStyle[] { FStyle.Bold } );
        }

        private static void InputTextIntoRow (int row, string docName, string docType, int qty, string comment)
        {
            InputTextIntoCell(row, 1, docName);
            InputTextIntoCell(row, 2, docType);
            InputTextIntoCell(row, 3, qty.ToString());
            InputTextIntoCell(row, 4, comment);
        }

        private static void InputTextIntoCell (int row, int col, string textInput)              // Inputs text into a specific cell
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

        private static void FormatCell(int row, int col, FStyle[] styles)                       // Changes formatting of a specific cell
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

        private static RunProperties SetFormatting(string fontStyle, int fontSize)              // Sets cell text formatting
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

        private static void MergeCellsInColumn (int col)
        {
            TableCell tCell;

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

            foreach(TableRow tRow in tab.Elements<TableRow>())
            {
                tCell = tRow.Elements<TableCell>().ElementAt(col);

                if(IsEmptyCell(tCell))
                {
                    tCell.Append(tcPropNext.CloneNode(true));
                }
                else
                {
                    tCell.Append(tcPropStart.CloneNode(true));
                }
            }
        }

        private static Boolean IsEmptyCell(TableCell tCell)
        {
            try 
            {
                Text txt = tCell.Elements<Paragraph>().First().Elements<Run>().First().Elements<Text>().First();
            }
            catch (Exception e)
            {
                return true;
            }
            return false;
        }
    }
}
