using CommunityToolkit.WinUI.UI.Controls;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using System;
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
                originalDataTable = ExcelHelper.LoadExcelIntoDataTable(filePath, true);
                DataTableHelper.ColumnNametoListBox(originalDataTable, columnOptionsListBox);
                AddActionsToListBox(columnOptionsListBox, excelDataGrid);
                DataTableHelper.FillDataGrid(originalDataTable, excelDataGrid);
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
                    ItemsSource = new ObservableCollection<string> { "Keep", "ColToNewRow", "Rename", "Delete" },
                    SelectedIndex = 0
                };
                TextBox renameTextBox = new TextBox { Width = 100, Visibility = Visibility.Collapsed };
                ComboBox ColToNewRowComboBox = new ComboBox { Width = 100, Visibility = Visibility.Collapsed };
                var viewModel = this;
                actionsComboBox.SelectionChanged += (s, e) =>
                {
                    switch (actionsComboBox.SelectedItem as string)
                    {
                        case "Rename":
                            renameTextBox.Visibility = Visibility.Visible;
                            ColToNewRowComboBox.Visibility = Visibility.Collapsed;
                            break;
                        case "ColToNewRow":
                            renameTextBox.Visibility = Visibility.Collapsed;
                            ColToNewRowComboBox.Visibility = Visibility.Visible;
                            ColToNewRowComboBox.ItemsSource = new ObservableCollection<string>(
                                listBox.Items.Cast<StackPanel>()
                                    .Select(panel => ((TextBlock)panel.Children[0]).Text)
                                    .Where(name => name != columnName.Text && !IsColumnDeleted(listBox, name))
                            );
                            break;
                        default:
                            renameTextBox.Visibility = Visibility.Collapsed;
                            ColToNewRowComboBox.Visibility = Visibility.Collapsed;
                            break;
                    }

                    // 모든 ComboBox의 상태를 확인하여 처리
                    viewModel.ProcessColumns(listBox, dataGrid);
                };

                renameTextBox.TextChanged += (s, e) => viewModel.ProcessColumns(listBox, dataGrid);
                ColToNewRowComboBox.SelectionChanged += (s, e) => viewModel.ProcessColumns(listBox, dataGrid);

                optionPanel.Children.Add(actionsComboBox);
                optionPanel.Children.Add(renameTextBox);
                optionPanel.Children.Add(ColToNewRowComboBox);

                listBox.Items.Add(optionPanel);
            }
        }

        private bool IsColumnDeleted(ListBox listBox, string columnName)
        {
            foreach (var panel in listBox.Items.Cast<StackPanel>())
            {
                var currentColumnName = ((TextBlock)panel.Children[0]).Text;
                var action = ((ComboBox)panel.Children[1]).SelectedItem as string;

                if (currentColumnName == columnName && action == "Delete")
                {
                    return true;
                }
            }
            return false;
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
                    var renameTextBox = (TextBox)panel.Children[2];
                    var ColToNewRowComboBox = (ComboBox)panel.Children[3];

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
                        case "Delete":
                            if (modifiedDataTable.Columns.Contains(columnName))
                            {
                                modifiedDataTable.Columns.Remove(columnName);
                            }
                            break;
                        case "Rename":
                            var newColumnName = renameTextBox.Text;
                            if (!string.IsNullOrEmpty(newColumnName) && modifiedDataTable.Columns.Contains(columnName))
                            {
                                modifiedDataTable.Columns[columnName].ColumnName = newColumnName;
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
                DataTable newDataTable = dataTable.Clone();

                foreach (DataRow row in dataTable.Rows)
                {
                    // 원본 행을 추가
                    newDataTable.ImportRow(row);

                    // 피벗 행을 추가
                    DataRow pivotRow = newDataTable.NewRow();
                    pivotRow.ItemArray = row.ItemArray.Clone() as object[];
                    pivotRow[targetColumnName] = row[sourceColumnName];
                    newDataTable.Rows.Add(pivotRow);
                }

                // 피벗 컬럼 삭제
                newDataTable.Columns.Remove(sourceColumnName);

                return newDataTable;
            }

            return dataTable;
        }
    }
}