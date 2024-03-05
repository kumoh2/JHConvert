using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using System;
using System.Collections.ObjectModel;
using System.Data;
using Windows.Storage.Pickers;
using Windows.Storage;
using WinRT.Interop;
using System.Linq;

namespace jhconvert
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            this.InitializeComponent();
        }

        private async void LoadExcel_Click(object sender, RoutedEventArgs e)
        {
            var picker = new FileOpenPicker();
            IntPtr hwnd = WindowNative.GetWindowHandle(this);
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
                var filePath = file.Path;
                bool firstRowAsColumnNames = false; // Assuming you have a checkbox or some UI to determine this
                DataTable dataTable = ExcelHelper.LoadExcelIntoDataTable(filePath, firstRowAsColumnNames);
                DataGridHelper.FillDataGrid(dataTable, ExcelDataGrid);
                UpdateColumnOptions(dataTable);
            }
        }

        private void UpdateColumnOptions(DataTable dataTable)
        {
            ColumnOptionsListBox.Items.Clear();
            for (int i = 0; i < dataTable.Columns.Count; i++)
            {
                DataColumn column = dataTable.Columns[i];

                StackPanel optionPanel = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 5, 0, 5) };

                // 순번을 컬럼 이름 앞에 추가
                TextBlock columnName = new TextBlock { Text = $"{i + 1}. {column.ColumnName}", Width = 150 };
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
                            pivotComboBox.ItemsSource = new ObservableCollection<string>(dataTable.Columns.Cast<DataColumn>().Select(c => c.ColumnName));
                            break;
                        default:
                            renameTextBox.Visibility = Visibility.Collapsed;
                            pivotComboBox.Visibility = Visibility.Collapsed;
                            break;
                    }
                };

                optionPanel.Children.Add(columnName);
                optionPanel.Children.Add(actionsComboBox);
                optionPanel.Children.Add(renameTextBox);
                optionPanel.Children.Add(pivotComboBox);

                ColumnOptionsListBox.Items.Add(optionPanel);
            }
        }
    }
}
