using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Linq;

namespace ExcelStyleVisualizer
{
    class Program
    {
        static void Main(string[] args)
        {
            string filePath = @"C:\temp\NeuesExcel.xlsx";

            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, true))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                var stylesPart = workbookPart.WorkbookStylesPart;
                var cellFormats = stylesPart.Stylesheet.CellFormats;

                // Neues Arbeitsblatt für die StyleIndex-Visualisierung erstellen
                WorksheetPart newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                newWorksheetPart.Worksheet = new Worksheet(new SheetData());
                Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>();
                string relationshipId = spreadsheetDocument.WorkbookPart.GetIdOfPart(newWorksheetPart);

                uint sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
                Sheet newSheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = "StyleIndex Visualization" };
                sheets.Append(newSheet);

                SheetData newSheetData = newWorksheetPart.Worksheet.GetFirstChild<SheetData>();

                uint currentRowIndex = 1;

                // Jedes CellFormat in einer Zelle visualisieren
                for (uint i = 0; i < cellFormats.Count(); i++)
                {
                    Row newRow = new Row() { RowIndex = currentRowIndex };
                    newSheetData.Append(newRow);

                    // StyleIndex in Zelle A darstellen
                    Cell indexCell = new Cell() { CellReference = "A" + currentRowIndex, CellValue = new CellValue($"Style {i}"), DataType = CellValues.String };
                    newRow.Append(indexCell);

                    // Visualisierung der StyleIndex-Formatierung in Zelle B
                    Cell styleCell = new Cell() { CellReference = "B" + currentRowIndex, StyleIndex = i, CellValue = new CellValue(" "), DataType = CellValues.String };
                    newRow.Append(styleCell);

                    // Eine leere Zeile einfügen
                    currentRowIndex += 2; // Zählt zwei nach unten, um eine leere Zeile einzufügen
                }

                newWorksheetPart.Worksheet.Save();
            }

            Console.WriteLine("StyleIndex-Visualisierung mit Leerzeilen erstellt.");
        }
    }
}
