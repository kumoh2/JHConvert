using CommunityToolkit.WinUI.UI.Controls;
using Microsoft.UI.Xaml.Data;
using Microsoft.UI.Xaml;
using System.Collections.ObjectModel;
using System.Data;
using Microsoft.UI.Xaml.Controls;

public static class DataTableHelper
{
    public static void FillDataGrid(DataTable table, DataGrid grid)
    {
        grid.Columns.Clear();
        grid.AutoGenerateColumns = false;
        for (int i = 0; i < table.Columns.Count; i++)
        {
            grid.Columns.Add(new DataGridTextColumn()
            {
                Header = table.Columns[i].ColumnName,
                Binding = new Binding { Path = new PropertyPath("[" + i.ToString() + "]") }
            });
        }

        var collection = new ObservableCollection<object>();
        foreach (DataRow row in table.Rows)
        {
            collection.Add(row.ItemArray);
        }

        grid.ItemsSource = collection;
    }

    public static void ColumnNametoListBox(DataTable dataTable, ListBox listBox)
    {
        listBox.Items.Clear();
        for (int i = 0; i < dataTable.Columns.Count; i++)
        {
            DataColumn column = dataTable.Columns[i];

            // 컬럼 이름 추가
            TextBlock columnName = new TextBlock { Text = column.ColumnName, Width = 100 };
            listBox.Items.Add(columnName);
        }
    }
}
