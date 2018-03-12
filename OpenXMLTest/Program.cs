using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLTest
{
    class Program
    {
        static void ReadXLSX(string fileName)
        {
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fileName, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
                SharedStringTable sharedStringTable = workbookPart.SharedStringTablePart.SharedStringTable;

                foreach (Row row in sheetData.Elements<Row>())
                {
                    foreach (Cell cell in row.Elements<Cell>())
                    {
                        if (cell.DataType != null && cell.DataType == CellValues.SharedString)
                        {
                            Console.Write(sharedStringTable.ElementAt(int.Parse(cell.InnerText)).InnerText);
                        }
                        else if (cell.CellValue != null)
                        {
                            Console.Write(cell.CellValue.Text);
                        }

                        Console.Write("\t");
                    }

                    Console.WriteLine();
                }
            }
        }

        static void Main(string[] args)
        {
            // First step: install DocumentFormat.OpenXml (by Microsoft) NuGet package

            try
            {
                ReadXLSX("test1.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
