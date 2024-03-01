using System;
using System.Data;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;

public static class ExcelHelper
{
    public static DataTable LoadExcelIntoDataTable(string filePath, bool firstRowAsColumnNames)
    {
        DataTable dataTable = new DataTable();

        using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
        {
            WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
            Sheet firstSheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault();
            WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(firstSheet.Id);
            SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

            var rows = sheetData.Elements<Row>().ToList();
            int columnCount = rows.First().Elements<Cell>().Count();

            if (firstRowAsColumnNames)
            {
                var headerRow = rows.First();
                foreach (Cell cell in headerRow.Elements<Cell>())
                {
                    dataTable.Columns.Add(GetCellValue(spreadsheetDocument, cell));
                }
                rows = rows.Skip(1).ToList();
            }
            else
            {
                for (int i = 0; i < columnCount; i++)
                {
                    dataTable.Columns.Add($"Column{i + 1}");
                }
            }

            dataTable.BeginLoadData();
            foreach (Row row in rows)
            {
                DataRow dataRow = dataTable.NewRow();
                int columnIndex = 0;

                foreach (Cell cell in row.Elements<Cell>())
                {
                    if (columnIndex < dataTable.Columns.Count)
                    {
                        dataRow[columnIndex] = GetCellValue(spreadsheetDocument, cell);
                    }
                    columnIndex++;
                }

                dataTable.Rows.Add(dataRow);
            }
            dataTable.EndLoadData();
        }

        return dataTable;
    }

    private static string GetCellValue(SpreadsheetDocument document, Cell cell)
    {
        string value = cell.CellValue?.InnerText;

        if (cell.DataType != null)
        {
            switch (cell.DataType.Value.ToString())
            {
                case "s":
                    var stringTable = document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                    value = stringTable.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
                    break;
                case "b":
                    value = value == "1" ? "TRUE" : "FALSE";
                    break;
            }
        }

        return value ?? string.Empty;
    }
}