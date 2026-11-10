using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Drawing;

class Program
{
    static void Main()
    {
        string filePath = "C:/Documenti/ProgettoRBG.xlsx";

        Console.WriteLine("Inserire il numero di colori da inserire: ");
        int n = Convert.ToInt16(Console.ReadLine());
        // 1. Chiedi i valori RGB per il colore
        Console.WriteLine("\nInserisci i valori per il colore (da 0 a 255):");

        Console.Write("Rosso (R): ");
        int r = int.Parse(Console.ReadLine());

        Console.Write("Verde (G): ");
        int g = int.Parse(Console.ReadLine());

        Console.Write("Blu (B): ");
        int b = int.Parse(Console.ReadLine());

        // 2. Genera il file Excel con ClosedXML
        using (var workbook = new XLWorkbook())
        {


            // Applica il colore inserito dall'utente alla cella scelta
            var worksheet = workbook.Worksheets.Add("Dati");

            var coloreUtente = XLColor.FromArgb(r, g, b);
            worksheet.Cell(1, 1).Style.Fill.BackgroundColor = coloreUtente;


            // scrive i valori RGB dentro la cella stessa
            worksheet.Cell(1, 1).Value = $"RGB({r},{g},{b})";
            worksheet.Columns().AdjustToContents();
            worksheet.Column(1).Width = 160;
            worksheet.Row(1).Height = 80;


            // Salva il file
            string nomeFile = "C:/Documenti/ProgettoRBG.xlsx";
            workbook.SaveAs(nomeFile);

            Console.WriteLine($"\nFile '{nomeFile}' creato con successo! Cella  colorata.");
        }
    }
}
