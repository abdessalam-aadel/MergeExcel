using System;
using System.Data;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.IO;
using System.Diagnostics;

namespace MergeExcel
{
    public partial class frmMain : Form
    {
        // Array of Excel Files found in Folder
        string[] XLSfiles;

        // Store slected path of Folder browser dialog in variable
        string selected_path;

        // Create fileCount to counting number of Excel files found
        int fileCount = 0;

        Excel.Application excelApplication = null;
        Excel.Workbook excelWorkBook = null;
        object paramMissing = Type.Missing;

        public frmMain()
        {
            InitializeComponent();
        }

        // Handle Methode Search in all Sub-Directory and Get all Excel files found,
        // and bring out to the string array
        private int SearchXLSFiles(string path, out string[] XLSfiles)
        {
            XLSfiles = Directory
                        .GetFiles(path, "*.*", SearchOption.AllDirectories)
                        .Where(s => s.ToLower().EndsWith(".xls") || s.ToLower().EndsWith(".xlsx"))
                        .ToArray();
            return XLSfiles.Length;
        }

        // Methode Write exceptions into log file
        static void LogException(string logFilePath, string filePath, Exception ex)
        {
            using (StreamWriter writer = new StreamWriter(logFilePath, true))
            {
                string filename = Path.GetFileNameWithoutExtension(filePath);
                writer.WriteLine($"{filename} : {ex.Message}");
            }
        }

        public void MergeExcelFiles(string[] filePaths, string destinationFile)
        {
            // Start Excel application
            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = false; // Make Excel invisible during processing
            // Disable display alerts to avoid the clipboard warning message
            excelApp.DisplayAlerts = false;

            // Create a new workbook to hold the merged data
            Excel.Workbook destinationWorkbook = excelApp.Workbooks.Add();
            Excel.Worksheet destinationWorksheet = destinationWorkbook.Sheets[1];

            int currentRow = 1;
            // Log file
            string logFilePath = selected_path + @"\exceptions.log";
            // Delete the log file if it exists
            if (File.Exists(logFilePath))
            {
                File.Delete(logFilePath);
            }

            foreach (string filePath in filePaths)
            {
                try
                {
                    // Open the source workbook
                    Excel.Workbook sourceWorkbook = excelApp.Workbooks.Open(filePath);

                    foreach (Excel.Worksheet sourceWorksheet in sourceWorkbook.Sheets)
                    {
                        // Find the last row with data in the source worksheet
                        Excel.Range sourceRange = sourceWorksheet.UsedRange;
                        int lastRow = sourceRange.Rows.Count;

                        // Copy the range from the source worksheet
                        Excel.Range rangeToCopy = sourceWorksheet.Range["A1", sourceWorksheet.Cells[lastRow, sourceRange.Columns.Count]];
                        rangeToCopy.Copy();

                        // Paste the data into the destination worksheet
                        destinationWorksheet.Cells[currentRow, 1].PasteSpecial(Excel.XlPasteType.xlPasteValues);

                        // Update the current row position for the next sheet
                        currentRow += lastRow;

                        // Optionally, you can copy headers and add extra checks to ensure the right data is copied.
                    }

                    // Close the source workbook (no save)
                    sourceWorkbook.Close(false);
                }
                catch (Exception ex)
                {
                    // Write Exception into exceptions.log
                    LogException(logFilePath, filePath, ex);
                    continue;
                }
                
            }

            // Save the destination workbook
            destinationWorkbook.SaveAs(destinationFile);
            destinationWorkbook.Close(false);

            // Quit Excel application
            excelApp.Quit();
            Marshal.ReleaseComObject(excelApp);
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            if (XLSfiles == null || string.IsNullOrEmpty(TxtBoxLoad.Text))
            {
                labelErrorMessage.Text = "No source folder was selected, Please select one.";
                return;
            }

            else if (XLSfiles.Length == 0)
            {
                labelErrorMessage.Text = "No Excel file was found in the selected folder";
                return;
            }

            labelErrorMessage.Text = "";
            Cursor = Cursors.WaitCursor;
            labelInfo.Text = "Processing ...";
            labelErrorMessage.Text = "";

            string destinationFile = @"C:\Users\P50\Desktop\mergedfile.xlsx";
            try
            {
                MergeExcelFiles(XLSfiles, destinationFile);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            // Clear string array
            XLSfiles = null;
            Cursor = Cursors.Default;
            TxtBoxLoad.Text = "Chose your folder location ...";
            labelInfo.Text = "Done.";

        }

        private void btnLoad_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog FD = new FolderBrowserDialog();
            if (selected_path != null)
                FD.SelectedPath = selected_path;
            if (FD.ShowDialog() == DialogResult.OK)
            {
                string path = FD.SelectedPath;
                selected_path = path;
                TxtBoxLoad.Text = path;
                fileCount = SearchXLSFiles(path, out XLSfiles);
                labelInfo.Text = fileCount + " XLS files found.";
            }
        }

        // Start Method CloseWorkBook
        private void CloseWorkBook()
        {
            if (excelWorkBook != null)
            {
                // Close the workbook object.
                excelWorkBook.Close(false, paramMissing, paramMissing);
                excelWorkBook = null;
            }
        }

        // Start Method QuitExcel
        private void QuitExcel()
        {
            // Quit Excel and release the ApplicationClass object.
            if (excelApplication != null)
            {
                excelApplication.Quit();
                excelApplication = null;
            }

            // Force garbage collection.
            GC.Collect();
            // Wait for all finalizers to complete before continuing.
            // Without this call to GC.WaitForPendingFinalizers,
            // the worker loop below might execute at the same time
            // as the finalizers. 
            // With this call, the worker loop executes only after
            // all finalizers have been called.
            GC.WaitForPendingFinalizers();
            // Clear string array
            XLSfiles = null;
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            // Go to Github repository
            string url = "https://github.com/abdessalam-aadel/MergeExcel";

            // Open the URL in the default web browser
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = url,
                    UseShellExecute = true // Ensures the URL is opened in the default web browser
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}");
            }
        }

        private void frmMain_DragDrop(object sender, DragEventArgs e)
        {
            string path = ((string[])e.Data.GetData(DataFormats.FileDrop))[0];
            // Condition >> Drag Folder
            if (Directory.Exists(path))
            {
                TxtBoxLoad.Text = path;
                fileCount = SearchXLSFiles(path, out XLSfiles);
                selected_path = path;
                // Check the Empty Folder
                labelInfo.Text = fileCount == 0 ? "Your Folder is Empty." : fileCount + " XLS files found.";
                labelErrorMessage.Text = "";
            }
        }

        private void frmMain_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Copy;
            TxtBoxLoad.Text = "Chose your folder location ...";
            labelInfo.Text = "...";
            labelErrorMessage.Text = "";
            XLSfiles = null;
        }
    }
}
