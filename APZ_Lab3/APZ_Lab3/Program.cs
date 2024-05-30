using FileExporters;

internal class Program
{
    private static void Main(string[] args)
    {
        string saveWordPath = @"E:\Docs\ХАІ\АПЗ\testword.docx"; // Шлях до збереження фалу
        string saveExсelPath = @"E:\Docs\ХАІ\АПЗ\testexel.xlsx"; // Шлях до збереження фалу

        // Створення Word документа
        WordExporter.Export(saveWordPath);

        // Створення Excel документа
        ExcelExporter.Export(saveExсelPath);
    }
}