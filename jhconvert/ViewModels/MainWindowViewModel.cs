using CommunityToolkit.WinUI.UI.Controls;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using Windows.Storage;
using Windows.Storage.Pickers;
using WinRT.Interop;

namespace jhconvert.ViewModels
{
    public class MainWindowViewModel
    {
        private DataTable originalDataTable;

        public async Task LoadExcel(IntPtr hwnd, ListBox columnOptionsListBox, DataGrid excelDataGrid)
        {
            var picker = new FileOpenPicker();
            InitializeWithWindow.Initialize(picker, hwnd);

            picker.ViewMode = PickerViewMode.Thumbnail;
            picker.FileTypeFilter.Add("*");
            picker.FileTypeFilter.Add(".xlsx");
            picker.FileTypeFilter.Add(".xlsm");
            picker.FileTypeFilter.Add(".xlsb");
            picker.FileTypeFilter.Add(".xls");

            StorageFile file = await picker.PickSingleFileAsync();
            if (file != null)
            {
                // Excel load start
                var filePath = file.Path; 
                originalDataTable = await Task.Run(() => ExcelHelper.LoadExcelIntoDataTable(filePath, true));
                DataTableHelper.ColumnNametoListBox(originalDataTable, columnOptionsListBox);
                AddActionsToListBox(columnOptionsListBox, excelDataGrid);
                DataTableHelper.FillDataGrid(originalDataTable, excelDataGrid);
                //excelDataGrid.Source = originalDataTable.DefaultView;
            }
        }

        public void ExportExcel()
        {
            // Export Excel logic
        }

        public void AddActionsToListBox(ListBox listBox, DataGrid dataGrid)
        {
            var items = listBox.Items.Cast<TextBlock>().ToList();
            listBox.Items.Clear();

            foreach (var columnName in items)
            {
                StackPanel optionPanel = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 5, 0, 5) };

                // 기존 TextBlock을 StackPanel에 추가
                optionPanel.Children.Add(columnName);

                ComboBox actionsComboBox = new ComboBox
                {
                    Width = 100,
                    ItemsSource = new ObservableCollection<string> { "Keep", "ColToNewRow" },
                    SelectedIndex = 0
                };
                ComboBox ColToNewRowComboBox = new ComboBox { Width = 100, Visibility = Visibility.Collapsed };
                var viewModel = this;
                actionsComboBox.SelectionChanged += (s, e) =>
                {
                    switch (actionsComboBox.SelectedItem as string)
                    {
                        case "ColToNewRow":
                            ColToNewRowComboBox.Visibility = Visibility.Visible;
                            ColToNewRowComboBox.ItemsSource = new ObservableCollection<string>(
                                listBox.Items.Cast<StackPanel>()
                                    .Select(panel => ((TextBlock)panel.Children[0]).Text)
                                    .Where(name => name != columnName.Text)
                            );
                            break;
                        default:
                            ColToNewRowComboBox.Visibility = Visibility.Collapsed;
                            break;
                    }

                    if (actionsComboBox.SelectedItem as string == "ColToNewRow" && string.IsNullOrEmpty(ColToNewRowComboBox.SelectedItem as string))
                    {
                        return; // ColToNewRowComboBox가 비어있으면 이벤트 종료
                    }

                    viewModel.ProcessColumns(listBox, dataGrid);
                };

                ColToNewRowComboBox.SelectionChanged += (s, e) => viewModel.ProcessColumns(listBox, dataGrid);

                optionPanel.Children.Add(actionsComboBox);
                optionPanel.Children.Add(ColToNewRowComboBox);

                listBox.Items.Add(optionPanel);
            }
        }

        private void ProcessColumns(ListBox listBox, DataGrid dataGrid)
        {
            if (originalDataTable != null)
            {
                DataTable modifiedDataTable = originalDataTable.Copy();

                foreach (var panel in listBox.Items.Cast<StackPanel>())
                {
                    var columnName = ((TextBlock)panel.Children[0]).Text;
                    var action = ((ComboBox)panel.Children[1]).SelectedItem as string;
                    var ColToNewRowComboBox = (ComboBox)panel.Children[2];

                    switch (action)
                    {
                        case "Keep":
                            // 아무 작업도 하지 않음
                            break;
                        case "ColToNewRow":
                            var ColToNewRowColumnName = ColToNewRowComboBox.SelectedItem as string;
                            if (!string.IsNullOrEmpty(ColToNewRowColumnName))
                            {
                                modifiedDataTable = AddColToNewRowColumn(modifiedDataTable, columnName, ColToNewRowColumnName);
                            }
                            break;
                    }
                }

                DataTableHelper.FillDataGrid(modifiedDataTable, dataGrid);
            }
        }

        private DataTable AddColToNewRowColumn(DataTable dataTable, string sourceColumnName, string targetColumnName)
        {
            if (dataTable.Columns.Contains(sourceColumnName) && dataTable.Columns.Contains(targetColumnName))
            {
                // 새로운 DataTable을 생성하지 않고, 원본 DataTable에 추가합니다.
                List<DataRow> newRows = new List<DataRow>();

                foreach (DataRow row in dataTable.Rows)
                {
                    // 원본 행을 추가할 때는 그대로 둡니다.
                    DataRow newRow = dataTable.NewRow();
                    newRow.ItemArray = row.ItemArray.Clone() as object[];
                    newRows.Add(newRow);

                    // 피벗 행을 추가합니다.
                    DataRow pivotRow = dataTable.NewRow();
                    pivotRow.ItemArray = row.ItemArray.Clone() as object[];
                    pivotRow[targetColumnName] = row[sourceColumnName];
                    newRows.Add(pivotRow);
                }

                // 새로운 행들을 원본 DataTable에 추가합니다.
                foreach (var newRow in newRows)
                {
                    dataTable.Rows.Add(newRow);
                }

                // 원본 컬럼을 제거합니다.
                dataTable.Columns.Remove(sourceColumnName);

                return dataTable;
            }

            return dataTable;
        }
    }
}