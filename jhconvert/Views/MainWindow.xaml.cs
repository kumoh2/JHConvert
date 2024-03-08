using Microsoft.UI.Xaml;
using jhconvert.ViewModels;
using System;

namespace jhconvert.Views
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
            IntPtr hwnd = WinRT.Interop.WindowNative.GetWindowHandle(this);
            await ViewModel.LoadExcel(hwnd, ColumnOptionsListBox, ExcelDataGrid);
        }

        private void ExportExcel_Click(object sender, RoutedEventArgs e)
        {
            ViewModel.ExportExcel();
        }
    }
}
