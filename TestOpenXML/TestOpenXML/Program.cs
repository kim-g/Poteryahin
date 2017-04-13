using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Data;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace TestOpenXML
{
    class Program
    {
        static void Main(string[] args)
        {
            ReadToDataTable(@"d:\Visual Studio\Poteryahin\TableParser\In\Test.xlsx");
        }

        static void ReadExcelFileDOM(string fileName)
        {
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fileName, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
                string text;
                Console.WriteLine(sheetData.Elements<Row>().Count());
                foreach (Row r in sheetData.Elements<Row>())
                {
                    foreach (Cell c in r.Elements<Cell>())
                    {
                        Console.Write(c.GetType() + ": ");
                        text = c.CellValue == null ? "" : c.CellValue.Text;
                        Console.Write(text + "; ");
                    }
                }
                Console.WriteLine();
                Console.ReadKey();
            }
        }

        // The SAX approach.
        static void ReadExcelFileSAX(string fileName)
        {
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fileName, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();

                OpenXmlReader reader = OpenXmlReader.Create(worksheetPart);
                string text;
                while (reader.Read())
                {
                    text = reader.GetText();
                    Console.Write(text + " ");

                }
                Console.WriteLine();
                Console.ReadKey();
            }
        }

        static void ReadExcelFile(string fileName)
        {
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fileName, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();

                OpenXmlReader reader = OpenXmlReader.Create(worksheetPart);
                string text;
                while (reader.Read())
                {
                    if (reader.ElementType == typeof(CellValue))
                    {
                        text = reader.GetText();
                        Console.Write(text + " ");
                    }
                }
            }
            Console.WriteLine();
            Console.ReadKey();
        }

        public static List<string> ExcelReader(string fileLocation)
        {
            Excel.Application excel = new Excel.Application();
            Excel.Workbook workBook =
                excel.Workbooks.Open(fileLocation);
            workBook.SaveAs(
                fileLocation + ".csv",
                Excel.XlFileFormat.xlCSVWindows
            );
            workBook.Close(true);
            excel.Quit();
            List<string> valueList = null;
            using (StreamReader sr = new StreamReader(fileLocation + ".csv"))
            {
                string content = sr.ReadToEnd();
                valueList = new List<string>(
                    content.Split(
                        new string[] { "\r\n" },
                        StringSplitOptions.RemoveEmptyEntries
                    )
                );
            }
            //new FileInfo(fileLocation + ".csv").Delete();
            return valueList;
        }

        public static void ReadToDataTable(string FileName)
        {
            DataTable dt = new DataTable();

            using (SpreadsheetDocument spreadSheetDocument = SpreadsheetDocument.Open(FileName, false))
            {

                WorkbookPart workbookPart = spreadSheetDocument.WorkbookPart;
                IEnumerable<Sheet> sheets = spreadSheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
                string relationshipId = sheets.First().Id.Value;
                WorksheetPart worksheetPart = (WorksheetPart)spreadSheetDocument.WorkbookPart.GetPartById(relationshipId);
                Worksheet workSheet = worksheetPart.Worksheet;
                SheetData sheetData = workSheet.GetFirstChild<SheetData>();
                IEnumerable<Row> rows = sheetData.Descendants<Row>();

                foreach (Cell cell in rows.ElementAt(0))
                {
                    dt.Columns.Add(GetCellValue(spreadSheetDocument, cell));
                }

                foreach (Row row in rows) //this will also include your header row...
                {
                    DataRow tempRow = dt.NewRow();

                    for (int i = 0; i < row.Descendants<Cell>().Count(); i++)
                    {
                        tempRow[i] = GetCellValue(spreadSheetDocument, row.Descendants<Cell>().ElementAt(i));
                    }

                    dt.Rows.Add(tempRow);
                }

            }
            dt.Rows.RemoveAt(0); //...so i'm taking it out here.

            for (int i = 0; i < dt.Rows.Count; i++)
                for (int j = 0; j < dt.Rows[i].ItemArray.Count(); j++)
                    Console.Write(dt.Rows[i][j].ToString());

            Console.WriteLine();
            Console.ReadKey();

        }


        public static string GetCellValue(SpreadsheetDocument document, Cell cell)
        {
            SharedStringTablePart stringTablePart = document.WorkbookPart.SharedStringTablePart;
            string value = cell.CellValue == null? " " : cell.CellValue.InnerXml;

            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                return stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;
            }
            else
            {
                return value;
            }
        }

    }

 }
