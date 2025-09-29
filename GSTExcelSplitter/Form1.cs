using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace GSTExcelSplitter
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void btnSplitExcel_Click(object sender, EventArgs e)
        {
            //string sourceFile = @"C:\Temp\AirwallexData.xlsx"; // original file
            //string outputFolder = @"C:\Temp\Output\"; // folder to save split files

            // File picker for input Excel
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Excel Files|*.xlsx;*.xls";
            ofd.Title = "Select an Excel file to split";

            if (ofd.ShowDialog() != DialogResult.OK)
                return;

            string sourceFile = ofd.FileName;

            // Folder picker for output
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            fbd.Description = "Select a folder to save the split files";

            if (fbd.ShowDialog() != DialogResult.OK)
                return;

            string outputFolder = fbd.SelectedPath;

            int batchSize = 3600;

            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook sourceWorkbook = null;
            Excel.Worksheet sourceSheet = null;
            Excel.Range usedRange = null;
         
            try
            {
                excelApp.Visible = false;
                sourceWorkbook = excelApp.Workbooks.Open(sourceFile);
                sourceSheet = sourceWorkbook.Sheets[1];

                usedRange = sourceSheet.UsedRange;
                int rowCount = usedRange.Rows.Count;
                int colCount = usedRange.Columns.Count;

                // Pull entire sheet into a 2D array and read usedRange into memory (allData)
                object[,] allData = (object[,])usedRange.Value2;


                // Deduplicate using a HashSet
                var seen = new HashSet<string>();
                var dedupedRows = new List<object[]>();

                // Add header row first
                object[] header = new object[colCount];
                for (int c = 1; c <= colCount; c++)
                    header[c - 1] = allData[1, c];
                dedupedRows.Add(header);

                // Loop through rows starting from 2 (skip header)
                for (int r = 2; r <= rowCount; r++)
                {
                    string colA = allData[r, 1]?.ToString() ?? "";
                    string colB = allData[r, 2]?.ToString() ?? "";
                    string key = colA + "|" + colB;

                    if (seen.Add(key))
                    {
                        object[] row = new object[colCount];
                        for (int c = 1; c <= colCount; c++)
                            row[c - 1] = allData[r, c];
                        dedupedRows.Add(row);
                    }
                }

                // Write deduplicated rows back into Excel

                // Cleaer source sheet first
                sourceSheet.Cells.Clear();

                // Convert List<object[]> dedupedRows back into 2D array
                object[,] dedupedArray = new object[dedupedRows.Count, colCount];
                for (int r = 0; r < dedupedRows.Count; r++)
                {
                    for (int c = 0; c < colCount; c++)
                    {
                        dedupedArray[r, c] = dedupedRows[r][c];
                    }
                }

                // Write back to Excel
                Excel.Range writeRange = sourceSheet.Range[
                    sourceSheet.Cells[1, 1],
                    sourceSheet.Cells[dedupedRows.Count, colCount]
                    ];
                writeRange.Value2 = dedupedArray;

                rowCount = dedupedRows.Count;

                int counter = 1;
                int fileIndex = 1;

                while (counter <= rowCount)
                {
                    int endRow = Math.Min(counter + batchSize - 1, rowCount);
                    Console.WriteLine($"Processing rows {counter} to {endRow}");

                    // Create a new workbook for this batch
                    Excel.Workbook newWorkbook = excelApp.Workbooks.Add();
                    Excel.Worksheet newSheet = newWorkbook.Sheets[1];

                    // Copy Header Row for every batch file
                    Excel.Range headerRange = sourceSheet.Range[
                        sourceSheet.Cells[1, 1],
                        sourceSheet.Cells[1, colCount]
                        ];

                    Excel.Range headerDest = newSheet.Range[
                        newSheet.Cells[1, 1],
                        newSheet.Cells[1, colCount]
                        ];

                    headerDest.Value2 = headerRange.Value2;

                    // Method 1: Copy rows to new sheet
                    //for (int row = counter; row <= endRow; row++)
                    //{
                    //    for (int col = 1; col <= colCount; col++)
                    //    {
                    //        newSheet.Cells[row - counter + 1, col].Value =
                    //            sourceSheet.Cells[row, col].Value;
                    //    }
                    //}

                    // Method 2: Bulk copy the entire block instead of looping cell by cell
                    Excel.Range sourceRange = sourceSheet.Range[
                        sourceSheet.Cells[counter, 1],
                        sourceSheet.Cells[endRow, colCount]
                        ];

                    Excel.Range destRange = newSheet.Range[
                        newSheet.Cells[2, 1], // Start from row 2 because header is row 1
                        newSheet.Cells[endRow - counter + 2, colCount]
                        ];

                    object[,] data = sourceRange.Value2;
                    destRange.Resize[data.GetLength(0), data.GetLength(1)].Value2 = data;

                    // Save new workbook
                    string newFile = Path.Combine(outputFolder, $"Batch_{fileIndex}.xlsx");
                    newWorkbook.SaveAs(newFile);
                    newWorkbook.Close();

                    Console.WriteLine($"Saved {newFile}");

                    // Release batch ranges inside each iteration of while loop
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(headerRange);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(headerDest);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(sourceRange);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(destRange);
                    sourceRange = null;
                    destRange = null;

                    // Move to next batch
                    counter += batchSize;
                    fileIndex++;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
            finally
            {
                try
                {
                    // Close workbook if still open
                    if(sourceWorkbook != null)
                    {
                        sourceWorkbook?.Close(false);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(sourceWorkbook);
                        sourceWorkbook = null;
                    }

                    if(sourceSheet != null)
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(sourceSheet);
                        sourceSheet = null;
                    }

                    if(usedRange != null)
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(usedRange);
                        usedRange = null;
                    }

                    if(excelApp != null) 
                    {
                        excelApp.Quit();
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                        excelApp = null;
                    }
                }
                catch { }

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

            MessageBox.Show("Splitting completed!");

            // Force closes the form and app
            this.Close();
            Application.Exit();

        }
    }
}
