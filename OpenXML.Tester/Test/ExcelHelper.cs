using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXML.Tester.Test
{
    class ExcelHelper
    {
        public void Test()
        {
            using (var workbook = SpreadsheetDocument.Create(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Test.xlsx"), DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = workbook.AddWorkbookPart();

                workbook.WorkbookPart.Workbook = new Workbook();

                workbook.WorkbookPart.Workbook.Sheets = new Sheets();


                var sheetPart = workbook.WorkbookPart.AddNewPart<WorksheetPart>();
                var sheetData = new SheetData();
                sheetPart.Worksheet = new Worksheet(sheetData);

                Sheets sheets = workbook.WorkbookPart.Workbook.GetFirstChild<Sheets>();
                string relationshipId = workbook.WorkbookPart.GetIdOfPart(sheetPart);

                uint sheetId = 1;
                if (sheets.Elements<Sheet>().Count() > 0)
                {
                    sheetId =
                        sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
                }

                Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = "UnitTest" };
                sheets.Append(sheet);

                Row headerRow = new Row();

                var dataHelper = new DataHelper();
                var data = dataHelper.GetData();

                List<String> columns = new List<string>() { "Id", "albumId", "title", "Url", "thumbnailUrl" };
                foreach (var column in columns)
                {
                    Cell cell = new Cell();
                    cell.DataType = CellValues.String;
                    cell.CellValue = new CellValue(column);
                    headerRow.AppendChild(cell);
                }

                sheetData.AppendChild(headerRow);

                for (int i = 0; i < 1; i++)
                {
                    int k = 0;
                    foreach (RandomData item in data)
                    {
                        if (++k == 5)
                            break;

                        Row newRow = new Row();
                        AddCell(newRow, item.Id);
                        AddCell(newRow, item.Field1);
                        AddCell(newRow, item.Field2);
                        AddCell(newRow, item.Field3);
                        AddCell(newRow, item.Field4);
                        AddCell(newRow, item.Field5);
                        AddCell(newRow, item.Field6);
                        AddCell(newRow, item.Field7);
                        AddCell(newRow, item.Field8);
                        AddCell(newRow, item.Field9);
                        AddCell(newRow, item.Field10);
                        AddCell(newRow, item.Field11);
                        AddCell(newRow, item.Field12);
                        AddCell(newRow, item.Field13);
                        AddCell(newRow, item.Field14);
                        AddCell(newRow, item.Field15);
                        AddCell(newRow, item.Field16);
                        AddCell(newRow, item.Field17);
                        AddCell(newRow, item.Field18);
                        AddCell(newRow, item.Field19);

                        sheetData.AppendChild(newRow);
                    }
                }

                sheetPart.Worksheet.Save();
            }

        }

        private void AddCell(Row row, string cellValue)
        {
            Cell cell = new Cell();
            cell.DataType = CellValues.String;
            cell.CellValue = new CellValue(cellValue);
            row.AppendChild(cell);
        }

        private void AddCell(Row row, int cellValue)
        {
            Cell cell = new Cell();
            cell.DataType = CellValues.Number;
            cell.CellValue = new CellValue(cellValue.ToString());
            row.AppendChild(cell);
        }
    }

}

