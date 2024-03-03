using ClosedXML.Excel;
using System.Data;
using System.Linq;

public static class ExcelHelper
{
    public static DataTable LoadExcelIntoDataTable(string filePath, bool firstRowAsColumnNames)
    {
        DataTable dt = new DataTable();

        using (var workbook = new XLWorkbook(filePath))
        {
            var worksheet = workbook.Worksheets.First();
            var rows = worksheet.RangeUsed().RowsUsed().ToList();

            // 첫 번째 행을 컬럼명으로 사용
            var headerRow = rows.First();
            foreach (var cell in headerRow.Cells())
            {
                dt.Columns.Add(cell.GetValue<string>());
            }

            // 데이터 행 추가
            dt.BeginLoadData();
            foreach (var row in rows.Skip(firstRowAsColumnNames ? 1 : 0))
            {
                var dataRow = dt.NewRow();
                for (int i = 0; i < row.Cells().Count(); i++)
                {
                    dataRow[i] = row.Cell(i + 1).Value;
                }
                dt.Rows.Add(dataRow);
            }
            dt.EndLoadData();
        }

        return dt;
    }
}