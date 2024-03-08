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
        public ObservableCollection<StackPanel> ColumnOptions { get; set; }
        public ObservableCollection<object> ExcelData { get; set; }

        public MainWindowViewModel()
        {
            ColumnOptions = new ObservableCollection<StackPanel>();
            ExcelData = new ObservableCollection<object>();
        }

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
                DataTable dataTable = ExcelHelper.LoadExcelIntoDataTable(filePath, true);
                DataTableHelper.ColumnNametoListBox(dataTable, columnOptionsListBox);
                AddActionsToListBox(columnOptionsListBox);
                DataTableHelper.FillDataGrid(dataTable, excelDataGrid);
            }
        }

        public void ExportExcel()
        {
            // Export Excel logic
        }

        public static void AddActionsToListBox(ListBox listBox)
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
                    ItemsSource = new ObservableCollection<string> { "Keep", "Pivot", "Rename", "Delete" },
                    SelectedIndex = 0
                };
                TextBox renameTextBox = new TextBox { Width = 100, Visibility = Visibility.Collapsed };
                ComboBox pivotComboBox = new ComboBox { Width = 100, Visibility = Visibility.Collapsed };

                actionsComboBox.SelectionChanged += (s, e) =>
                {
                    switch (actionsComboBox.SelectedItem as string)
                    {
                        case "Rename":
                            renameTextBox.Visibility = Visibility.Visible;
                            pivotComboBox.Visibility = Visibility.Collapsed;
                            break;
                        case "Pivot":
                            renameTextBox.Visibility = Visibility.Collapsed;
                            pivotComboBox.Visibility = Visibility.Visible;
                            pivotComboBox.ItemsSource = new ObservableCollection<string>(listBox.Items.Cast<StackPanel>()
                                .Select(panel => ((TextBlock)panel.Children[0]).Text));
                            break;
                        default:
                            renameTextBox.Visibility = Visibility.Collapsed;
                            pivotComboBox.Visibility = Visibility.Collapsed;
                            break;
                    }
                };

                optionPanel.Children.Add(actionsComboBox);
                optionPanel.Children.Add(renameTextBox);
                optionPanel.Children.Add(pivotComboBox);

                listBox.Items.Add(optionPanel);
            }
        }
    }
}