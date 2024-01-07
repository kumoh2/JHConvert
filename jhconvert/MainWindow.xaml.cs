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
            PickFilesOutputTextBlock.Text = "엑셀을 불러오는 중";

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

                        // 원본 시트의 셀 너비 복사
                        int colCount = worksheet.ColumnsUsed().Count();
                        for (int i = 1; i <= colCount; i++)
                        {
                            newSheet.Column(i).Width = worksheet.Column(i).Width;
                        }

                        // 각 파일의 모든 행을 새 시트에 복사 (서식 포함)
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
                // 새 엑셀 파일 저장
                string filePath = "MergedFile.xlsx";
                newWorkbook.SaveAs(filePath);
                PickFilesOutputTextBlock.Text = output.ToString();
                // 파일 저장 후 실행
                try
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = filePath,
                        UseShellExecute = true // 파일과 연결된 기본 프로그램으로 실행
                    });
                }
                catch (Exception ex)
                {
                    // 파일을 여는 중 오류 처리
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
