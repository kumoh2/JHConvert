using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using System;
using System.Collections.ObjectModel;
using System.Data;
using Windows.Storage.Pickers;
using Windows.Storage;
using WinRT.Interop;
using System.Linq;
using jhconvert.ViewModels;

namespace jhconvert
{
    public partial class MainWindow : Window
    {
        public MainWindowViewModel ViewModel { get; }

        public MainWindow()
        {
            ViewModel = new MainWindowViewModel();
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
                //Excel load start
                var filePath = file.Path;
                DataTable dataTable = ExcelHelper.LoadExcelIntoDataTable(filePath, true);
                DataTableHelper.ColumnNametoListBox(dataTable, ColumnOptionsListBox);
                AddActionsToListBox(ColumnOptionsListBox);
                DataTableHelper.FillDataGrid(dataTable, ExcelDataGrid);
            }
        }

        private void ExportExcel_Click(object sender, RoutedEventArgs e)
        { 
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
