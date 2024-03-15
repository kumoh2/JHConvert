using CommunityToolkit.WinUI.UI.Controls;
using Microsoft.UI.Xaml.Data;
using Microsoft.UI.Xaml;
using System.Collections.ObjectModel;
using System.Data;
using Microsoft.UI.Xaml.Controls;
using System.Collections.Generic;

public static class DataTableHelper
{
    public static void FillDataGrid(DataTable table, DataGrid grid)
    {
        grid.Columns.Clear();
        grid.AutoGenerateColumns = false;

        // DataGridTextColumn 객체를 리스트에 미리 생성하여 저장
        var columns = new List<DataGridTextColumn>(table.Columns.Count);

        for (int i = 0; i < table.Columns.Count; i++)
        {
            // 캐싱 없이 Binding을 직접 사용
            columns.Add(new DataGridTextColumn
            {
                Header = table.Columns[i].ColumnName,
                Binding = new Binding { Path = new PropertyPath($"[{i}]") }
            });
        }

        foreach (var column in columns)
        {
            grid.Columns.Add(column);
        }

        var collection = new ObservableCollection<object>(table.AsEnumerable().Select(row => row.ItemArray));
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
