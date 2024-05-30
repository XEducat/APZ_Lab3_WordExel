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

            // Додайте необхідний текст і таблиці до документу Word

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
            }
        }
    }
}
