using CommunityToolkit.WinUI.UI.Controls;
using Microsoft.UI.Xaml.Data;
using Microsoft.UI.Xaml;
using System.Collections.ObjectModel;
using System.Data;

public static class DataGridHelper
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
}
