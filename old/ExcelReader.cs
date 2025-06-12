using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Linq;
using System.Collections.Generic;

public class ExcelReader
{
    public static List<List<string>> GetCellValuesFromRange(string fileName, string rangeName)
    {
        List<List<string>> cellValues = new List<List<string>> { };

        try
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
            {
                WorkbookPart workbookPart = document.WorkbookPart;
                var definedNames = workbookPart.Workbook.DefinedNames;
                if (definedNames != null)
                {
                    var gssInputRange = definedNames.Elements<DefinedName>().FirstOrDefault(dn => dn.Name == rangeName);
                    if (gssInputRange != null)
                    {
                        string[] range = gssInputRange.Text.Split('!')[1].Split(':');
                        string sheetName = gssInputRange.Text.Split('!')[0].Trim('\'');
                        Sheet sheet = workbookPart.Workbook.Sheets.Elements<Sheet>().FirstOrDefault(s => s.Name == sheetName);
                        if (sheet != null)
                        {
                            WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                            string startCellReference = range[0].Replace("$", ""); // Remove dollar signs
                            string endCellReference = (range.Length == 1) ? startCellReference : range[1].Replace("$", ""); // Remove dollar signs
                            Console.WriteLine($"Looking for cells from {startCellReference} to {endCellReference} in sheet: {sheetName}");

                            var cells = worksheetPart.Worksheet.Descendants<Cell>()
                                .Where(c => IsCellInRange(c.CellReference, startCellReference, endCellReference));

                            int numberOfColumns = GetColumnRowIndices(endCellReference).column - GetColumnRowIndices(startCellReference).column+1;
                            int numberOfRows = cells.Count() / numberOfColumns;

                            int i = 0;
                            List<string> row = new List<string>();
                            foreach (var cell in cells)
                            {
                                i++;
                                if (i <= numberOfColumns)
                                {
                                    row.Add(GetCellValue(workbookPart, cell));
                                }
                                else
                                {
                                    cellValues.Add(row); // Add the row to the main list
                                    i = 1;
                                    row = new List<string>();
                                    row.Add(GetCellValue(workbookPart, cell));
                                }
                            }
                            cellValues.Add(row); // Add the row to the main list


                            //foreach (var cell in cells)
                            //{
                            //    cellValues.Add(GetCellValue(workbookPart, cell));
                            //}
                        }
                        else
                        {
                            Console.WriteLine($"Sheet {sheetName} not found.");
                        }
                    }
                    else
                    {
                        Console.WriteLine($"Range {rangeName} not found.");
                    }
                }
                else
                {
                    Console.WriteLine("No defined names found.");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error reading {rangeName} range: {ex.Message}");
        }
        return cellValues;
    }

    private static bool IsCellInRange(string cellReference, string startCellReference, string endCellReference)
    {
        // Convert cell references to row and column indices
        (int startColumn, int startRow) = GetColumnRowIndices(startCellReference);
        (int endColumn, int endRow) = GetColumnRowIndices(endCellReference);
        (int cellColumn, int cellRow) = GetColumnRowIndices(cellReference);

        // Check if the cell is within the specified range
        return cellColumn >= startColumn && cellColumn <= endColumn && cellRow >= startRow && cellRow <= endRow;
    }

    private static (int column, int row) GetColumnRowIndices(string cellReference)
    {
        // Extract column letters and row numbers from the cell reference
        string columnLetters = new string(cellReference.Where(char.IsLetter).ToArray());
        string rowNumbers = new string(cellReference.Where(char.IsDigit).ToArray());

        // Convert column letters to column index (A=1, B=2, ..., Z=26, AA=27, etc.)
        int columnIndex = 0;
        foreach (char letter in columnLetters)
        {
            columnIndex = columnIndex * 26 + (letter - 'A' + 1);
        }

        // Convert row numbers to row index
        int rowIndex = int.Parse(rowNumbers);

        return (columnIndex, rowIndex);
    }

    private static string GetCellValue(WorkbookPart workbookPart, Cell cell)
    {
        string value = cell.InnerText;
        if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
        {
            var stringTable = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
            if (stringTable != null)
            {
                value = stringTable.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
            }
        }
        return value;
    }


    static void Main(string[] args)
    {
        var cellValues = GetCellValuesFromRange("C:\\InSync\\Lab\\files\\FakeTool_001.00.xlsx", "GSS_DATA");
        foreach (var rowValue in cellValues)
        {
            foreach (var cellValue in rowValue)
            {
                Console.Write($"cell value: {cellValue}");
            }
            Console.WriteLine();
        }
    }
}
