using ClosedXML.Excel;
using Microsoft.Win32;
using System;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls.Primitives;

namespace MagicA1
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Grid_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            // Allows the user to drag the window when clicking on the grid area
            if (e.ButtonState == System.Windows.Input.MouseButtonState.Pressed)
            {
                this.DragMove();
            }
        }

        private void MinimizeButton_Click(object sender, RoutedEventArgs e)
        {
            this.WindowState = WindowState.Minimized;
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close(); // Closes the window when the button is clicked
        }

        private void Border_DragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop, true))
            {
                string[] fileNames = (string[])e.Data.GetData(DataFormats.FileDrop, true);

                // Check if all dropped items are either directories or Excel files (.xlsx, .xlsm)
                bool allItemsValid = true;

                foreach (var fileName in fileNames)
                {
                    if (!Directory.Exists(fileName) &&
                        !IsExcelFile(fileName))
                    {
                        allItemsValid = false;
                        break;
                    }
                }

                if (allItemsValid)
                {
                    e.Effects = DragDropEffects.Copy;
                }
                else
                {
                    e.Effects = DragDropEffects.None;
                }
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }

            e.Handled = true;
        }

        private async void Border_Drop(object sender, DragEventArgs e)
        {
            try
            {
                if (e.Data.GetDataPresent(DataFormats.FileDrop, true))
                {
                    string[] fileNames = (string[])e.Data.GetData(DataFormats.FileDrop, true);

                    var tasks = new List<Task>();

                    foreach (var fileName in fileNames)
                    {
                        if (Directory.Exists(fileName))
                        {
                            // Process all Excel files in the directory (including subdirectories)
                            tasks.Add(ProcessAllExcelFilesInFolder(fileName));
                        }
                        else if (IsExcelFile(fileName))
                        {
                            // Process a single Excel file
                            tasks.Add(ProcessSingleExcelFile(fileName));
                        }
                    }

                    // Wait for all tasks to complete
                    await Task.WhenAll(tasks);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"{ex.Message}");
            }
        }

        private async Task ProcessSingleExcelFile(string filePath)
        {
            try
            {
                await Task.Run(() =>
                {
                    using (var workbook = new XLWorkbook(filePath))
                    {
                        var worksheet = workbook.Worksheet(1);
                        // Set A1 as the active cell
                        worksheet.Cell("A1").SetActive();

                        // Select only A1 and set the top-left cell to A1 in the view
                        worksheet.Range("A1").Select();

                        // Set the top-left cell of the visible view to A1
                        worksheet.SheetView.TopLeftCellAddress = worksheet.Cell("A1").Address;

                        // Save the changes to the file
                        workbook.Save();
                    }
                });

                // Update status in UI
                AddStatus(Path.GetFileName(filePath), "OK");
            }
            catch (Exception ex)
            {
                // Update status in UI
                AddStatus(Path.GetFileName(filePath), "NG");

                MessageBox.Show($"{ex.Message}");
            }
        }

        private async void SelectFolder_Click(object sender, RoutedEventArgs e)
        {
            var folderDialog = new OpenFolderDialog
            {
                Title = "Select Folder"
            };

            if (folderDialog.ShowDialog() == true)
            {
                string selectedFolder = folderDialog.FolderName;
                await ProcessAllExcelFilesInFolder(selectedFolder);
            }
        }

        private async Task ProcessAllExcelFilesInFolder(string folderPath)
        {
            try
            {
                var excelFiles = Directory.GetFiles(folderPath, "*.*", SearchOption.AllDirectories)
                    .Where(f => IsExcelFile(f));

                foreach (var filePath in excelFiles)
                {
                    await Task.Run(() =>
                    {
                        using (var workbook = new XLWorkbook(filePath))
                        {
                            var worksheet = workbook.Worksheet(1);
                            // Set A1 as the active cell
                            worksheet.Cell("A1").SetActive();

                            // Select only A1 and set the top-left cell to A1 in the view
                            worksheet.Range("A1").Select();

                            // Set the top-left cell of the visible view to A1
                            worksheet.SheetView.TopLeftCellAddress = worksheet.Cell("A1").Address;

                            // Save the changes to the file
                            workbook.Save();
                        }
                    });
                }

                // Update status in UI
                AddStatus(folderPath, "OK");
            }
            catch (Exception ex)
            {
                // Update status in UI
                AddStatus(folderPath, "NG");
                MessageBox.Show($"{ex.Message}");
            }
        }

        private bool IsExcelFile(string fileName)
        {
            var extension = Path.GetExtension(fileName).ToLower();
            return extension == ".xlsx" || extension == ".xlsm";
        }

#pragma warning disable CA1416 // Suppress platform-specific API warning
        private void AddStatus(string itemName, string status)
        {
            Dispatcher.Invoke(() =>
            {
                var duration = 3;
                Statusbar.MessageQueue?.Enqueue(
                   $"{itemName} - {status}",
                    null,
                    null,
                    null,
                    false,
                    true,
                    TimeSpan.FromSeconds(duration));
            });
        }
#pragma warning restore CA1416
    }
}
