using Microsoft.UI.Xaml;
using System;
using System.Data;
using Windows.Storage.Pickers;
using Windows.Storage;
using WinRT.Interop;

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
                bool firstRowAsColumnNames = FirstRowAsHeader.IsChecked ?? false;
                DataTable dataTable = ExcelHelper.LoadExcelIntoDataTable(filePath, firstRowAsColumnNames);
                DataGridHelper.FillDataGrid(dataTable, ExcelDataGrid);
            }
        }     
    }
}
