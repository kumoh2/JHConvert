using System.Data;
using ClosedXML.Excel;

namespace WinFormsApp4
{
    public partial class Form1 : Form
    {
        private DataTable originalDataTable;

        public Form1()
        {
            InitializeComponent();
            checkedListBox1.ItemCheck += CheckedListBox1_ItemCheck;
            dataGridView1.RowPostPaint += dataGridView1_RowPostPaint;
            textBox1.KeyPress += textBox1_KeyPress;
            textBox1.TextChanged += textBox1_TextChanged;
            comboBox1.SelectedIndexChanged += ComboBox1_SelectedIndexChanged;
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            // 숫자와 백스페이스만 허용
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            // 빈 문자열이 아니어야 함
            if (string.IsNullOrEmpty(textBox1.Text))
            {
                textBox1.Text = "1";
                textBox1.SelectionStart = textBox1.Text.Length; // 커서를 끝으로 이동
            }
        }

        private void dataGridView1_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            using (SolidBrush b = new SolidBrush(dataGridView1.RowHeadersDefaultCellStyle.ForeColor))
            {
                e.Graphics.DrawString((e.RowIndex + 1).ToString(), e.InheritedRowStyle.Font, b, e.RowBounds.Location.X + 20, e.RowBounds.Location.Y + 4);
            }
        }

        private async void button1_Click(object sender, EventArgs e)
        {
            // 초기화
            ResetForm();

            using (var openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files|*.xls;*.xlsx;*.xlsm;*.xlsb|All Files|*.*"
            })
            {
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    var filePath = openFileDialog.FileName;
                    originalDataTable = await Task.Run(() => LoadExcelIntoDataTable(filePath, true));
                    PopulateColumnOptions(originalDataTable);
                    dataGridView1.DataSource = originalDataTable;
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            using (var saveFileDialog = new SaveFileDialog
            {
                Filter = "Excel Files|*.xlsx;*.xlsm;*.xlsb|All Files|*.*",
                FileName = "Export.xlsx"
            })
            {
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    var filePath = saveFileDialog.FileName;
                    int rowsPerFile = int.Parse(textBox1.Text);
                    ExportDataGridViewToExcel(dataGridView1, filePath, rowsPerFile);
                    MessageBox.Show("Data exported successfully!", "Export Excel", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        private DataTable LoadExcelIntoDataTable(string filePath, bool firstRowAsColumnNames)
        {
            var dt = new DataTable();

            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheets.First();
                var rows = worksheet.RangeUsed().RowsUsed();

                var headerRow = rows.First();
                var columnNames = new HashSet<string>();
                foreach (var cell in headerRow.Cells())
                {
                    // 줄바꿈 문자 제거
                    string columnName = cell.GetString().Replace("\n", "").Replace("\r", "");

                    // 중복된 컬럼 이름 처리
                    if (columnNames.Contains(columnName))
                    {
                        int suffix = 2; // 2부터 시작
                        string newColumnName;
                        do
                        {
                            newColumnName = $"{columnName} #{suffix++}";
                        } while (columnNames.Contains(newColumnName));
                        columnName = newColumnName;
                    }
                    columnNames.Add(columnName);
                    dt.Columns.Add(columnName);
                }

                dt.BeginLoadData();
                foreach (var row in rows.Skip(firstRowAsColumnNames ? 1 : 0))
                {
                    var dataRow = dt.NewRow();
                    for (int i = 0; i < row.Cells().Count(); i++)
                    {
                        dataRow[i] = row.Cell(i + 1).GetString();
                    }
                    dt.Rows.Add(dataRow);
                }
                dt.EndLoadData();
            }

            return dt;
        }

        private void ExportDataGridViewToExcel(DataGridView dgv, string baseFilePath, int rowsPerFile)
        {
            int fileIndex = 1;
            int totalRows = dgv.Rows.Count;
            string directory = System.IO.Path.GetDirectoryName(baseFilePath);
            string fileNameWithoutExtension = System.IO.Path.GetFileNameWithoutExtension(baseFilePath);
            string extension = System.IO.Path.GetExtension(baseFilePath);

            for (int startRow = 0; startRow < totalRows; startRow += rowsPerFile)
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Sheet1");

                    // 열 헤더 추가
                    AddHeadersToWorksheet(dgv, worksheet);

                    // 행 데이터 추가
                    bool hasData = AddRowsToWorksheet(dgv, worksheet, startRow, rowsPerFile);

                    if (hasData)
                    {
                        string filePath = $"{directory}\\{fileNameWithoutExtension}_{fileIndex++}{extension}";
                        workbook.SaveAs(filePath);
                    }
                }
            }
        }

        private void AddHeadersToWorksheet(DataGridView dgv, IXLWorksheet worksheet)
        {
            for (int i = 0; i < dgv.Columns.Count; i++)
            {
                worksheet.Cell(1, i + 1).Value = dgv.Columns[i].HeaderText;
            }
        }

        private bool AddRowsToWorksheet(DataGridView dgv, IXLWorksheet worksheet, int startRow, int rowsPerFile)
        {
            bool hasData = false;
            for (int i = 0; i < rowsPerFile && (startRow + i) < dgv.Rows.Count; i++)
            {
                hasData = true;
                for (int j = 0; j < dgv.Columns.Count; j++)
                {
                    worksheet.Cell(i + 2, j + 1).Value = dgv.Rows[startRow + i].Cells[j].Value?.ToString();
                }
            }
            return hasData;
        }

        private void PopulateColumnOptions(DataTable dataTable)
        {
            checkedListBox1.Items.Clear();
            comboBox1.Items.Clear();
            foreach (DataColumn column in dataTable.Columns)
            {
                // 줄바꿈 문자 제거
                string columnName = column.ColumnName.Replace("\n", "").Replace("\r", "");
                checkedListBox1.Items.Add(columnName);
                comboBox1.Items.Add(columnName);
            }
        }

        private void ComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            checkedListBox1.Items.Clear();

            foreach (DataColumn column in originalDataTable.Columns)
            {
                // 줄바꿈 문자 제거
                string columnName = column.ColumnName.Replace("\n", "").Replace("\r", "");
                checkedListBox1.Items.Add(columnName);
            }

            dataGridView1.DataSource = null;
            dataGridView1.DataSource = originalDataTable;
        }

        private void CheckedListBox1_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            var selectedColumn = comboBox1.SelectedItem.ToString();
            var checkedColumns = checkedListBox1.CheckedItems.Cast<string>().ToList();

            if (e.NewValue == CheckState.Checked)
            {
                checkedColumns.Add(checkedListBox1.Items[e.Index].ToString());
            }
            else if (e.NewValue == CheckState.Unchecked)
            {
                checkedColumns.Remove(checkedListBox1.Items[e.Index].ToString());
            }

            DataTable updatedDataTable = new DataTable();

            // 새 데이터 테이블에 원본 컬럼 추가
            updatedDataTable.Columns.Add(selectedColumn);
            foreach (DataColumn column in originalDataTable.Columns)
            {
                if (column.ColumnName != selectedColumn && !checkedColumns.Contains(column.ColumnName))
                {
                    updatedDataTable.Columns.Add(column.ColumnName);
                }
            }

            foreach (DataRow row in originalDataTable.Rows)
            {
                foreach (var column in checkedColumns)
                {
                    DataRow newRow = updatedDataTable.NewRow();
                    newRow[selectedColumn] = row[column];
                    foreach (DataColumn originalColumn in originalDataTable.Columns)
                    {
                        if (originalColumn.ColumnName != selectedColumn && !checkedColumns.Contains(originalColumn.ColumnName))
                        {
                            newRow[originalColumn.ColumnName] = row[originalColumn.ColumnName];
                        }
                    }
                    updatedDataTable.Rows.Add(newRow);
                }
            }

            dataGridView1.DataSource = updatedDataTable;
        }

        private void ResetForm()
        {
            // 원본 데이터 테이블 초기화
            originalDataTable = null;
            dataGridView1.DataSource = null;

            // UI 초기화
            checkedListBox1.Items.Clear();
            comboBox1.Items.Clear();
            textBox1.Text = "99999999";
        }
    }
}
