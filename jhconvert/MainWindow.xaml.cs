using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.UI.Xaml;
using Windows.Storage;
using Windows.Storage.Pickers;
using ClosedXML.Excel;
using System.IO;
using System.Diagnostics;
using System.Linq;

// To learn more about WinUI, the WinUI project structure,
// and more about our project templates, see: http://aka.ms/winui-project-info.

namespace App10
{
    /// <summary>
    /// An empty window that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class MainWindow : Window
    {
        public MainWindow()
        {
            this.InitializeComponent();
        }

        private async void PickFilesButton_Click(object sender, RoutedEventArgs e)
        {
            // Clear previous returned file name, if it exists, between iterations of this scenario
            PickFilesOutputTextBlock.Text = "������ �ҷ����� ��";

            // Create a file picker
            var openPicker = new FileOpenPicker();
            var hWnd = WinRT.Interop.WindowNative.GetWindowHandle(this);
            WinRT.Interop.InitializeWithWindow.Initialize(openPicker, hWnd);

            // Set options for your file picker
            openPicker.ViewMode = PickerViewMode.List;
            openPicker.SuggestedStartLocation = PickerLocationId.DocumentsLibrary;
            //openPicker.FileTypeFilter.Add("*");
            openPicker.FileTypeFilter.Add(".xlsx");

            // Open the picker for the user to pick a file
            IReadOnlyList<StorageFile> files = await openPicker.PickMultipleFilesAsync();
            if (files.Count > 0)
            {
                var newWorkbook = new XLWorkbook();
                var newSheet = newWorkbook.Worksheets.Add("Merged Data");
                int currentRow = 1;
                StringBuilder output = new StringBuilder("Picked files:\n");

                foreach (StorageFile file in files)
                {
                    using (var stream = await file.OpenStreamForReadAsync())
                    {
                        var workbook = new XLWorkbook(stream);
                        var worksheet = workbook.Worksheet(1);

                        // ���� ��Ʈ�� �� �ʺ� ����
                        int colCount = worksheet.ColumnsUsed().Count();
                        for (int i = 1; i <= colCount; i++)
                        {
                            newSheet.Column(i).Width = worksheet.Column(i).Width;
                        }

                        // �� ������ ��� ���� �� ��Ʈ�� ���� (���� ����)
                        foreach (var row in worksheet.RangeUsed().Rows())
                        {
                            foreach (var cell in row.Cells())
                            {
                                var newCell = newSheet.Cell(currentRow, cell.Address.ColumnNumber);
                                newCell.Value = cell.Value;
                                newCell.Style = cell.Style;
                            }
                            currentRow++;
                        }
                    }
                    output.Append(file.Name + "\n");
                }
                // �� ���� ���� ����
                string filePath = "MergedFile.xlsx";
                newWorkbook.SaveAs(filePath);
                PickFilesOutputTextBlock.Text = output.ToString();
                // ���� ���� �� ����
                try
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = filePath,
                        UseShellExecute = true // ���ϰ� ����� �⺻ ���α׷����� ����
                    });
                }
                catch (Exception ex)
                {
                    // ������ ���� �� ���� ó��
                    Debug.WriteLine("Cannot open file: " + ex.Message);
                }
            }
            else
            {
                PickFilesOutputTextBlock.Text = "Operation cancelled.";
            }
        }
    }
}
