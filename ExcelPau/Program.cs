using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Data;

class Program
{
    public static void Main(string[] args)
    {
        SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(@"C:\Users\Izagakhmaevra\Desktop\Excel\database.xlsx", SpreadsheetDocumentType.Workbook);

        WorkbookPart workbookPart = spreadsheetDocument.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();

        workbookPart.Workbook.AppendChild(new Sheets());

        WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet(new SheetData());

        Sheet sheet = new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = 1,
            Name = "TestExcel"
        };

        spreadsheetDocument.WorkbookPart.Workbook.Sheets.Append(sheet);

        Row row = new Row();

        SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
        sheetData.Append(row);

        Cell cell = new Cell
        {
            CellReference = "A1",
            CellValue = new CellValue("Excel init"),
            DataType = CellValues.String,
        };


        row.Append(cell);

        workbookPart.Workbook.Save();
        spreadsheetDocument.Close();

    }
}
