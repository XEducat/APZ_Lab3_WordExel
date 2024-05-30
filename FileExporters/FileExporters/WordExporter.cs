using System.Runtime.InteropServices;
using System;
using Word = Microsoft.Office.Interop.Word;

namespace FileExporters
{
    public class WordExporter
    {
        public static void Export(string path)
        {
            Word.Application wdApp = new Word.Application();
            Word.Document doc = wdApp.Documents.Add();
            Word.Paragraph p = doc.Paragraphs.Add();

            p.Range.Text = "„Безнадійно — це коли на кришку труни падає земля. Решту можна виправити.“ — Джейсон Стетхем";
            p.Range.InsertParagraphAfter();

            Word.Table table = doc.Tables.Add(p.Range, 3, 3);

            table.Cell(1, 1).Range.Text = "Заголовок 1";
            table.Cell(1, 2).Range.Text = "Заголовок 2";
            table.Cell(1, 3).Range.Text = "Заголовок 3";

            table.Cell(2, 1).Range.Text = "Комірка 1";
            table.Cell(2, 2).Range.Text = "Комірка 2";
            table.Cell(2, 3).Range.Text = "Комірка 3";

            table.Cell(3, 1).Range.Text = "Комірка 4";
            table.Cell(3, 2).Range.Text = "Комірка 5";
            table.Cell(3, 3).Range.Text = "Комірка 6";

            table.Borders.Enable = 1;
            table.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            table.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;

            try
            {
                doc.SaveAs(path);
                Console.WriteLine("Word документ успiшно створений");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Виникла помилка при збереженнi документу Word за шляхом {path}");
            }
            finally
            {
                doc.Close();
                wdApp.Quit();
                Marshal.ReleaseComObject(doc);
                Marshal.ReleaseComObject(wdApp);
                Marshal.ReleaseComObject(p);
            }
        }
    }
}
