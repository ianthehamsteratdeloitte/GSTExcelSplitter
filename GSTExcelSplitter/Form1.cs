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
            string sourceFile = @"C:\Temp\AirwallexData.xlsx"; // original file
            string outputFolder = @"C:\Temp\Output\"; // folder to save split files
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

                int counter = 1;
                int fileIndex = 1;

                while (counter <= rowCount)
                {
                    int endRow = Math.Min(counter + batchSize - 1, rowCount);
                    Console.WriteLine($"Processing rows {counter} to {endRow}");

                    // Create a new workbook for this batch
                    Excel.Workbook newWorkbook = excelApp.Workbooks.Add();
                    Excel.Worksheet newSheet = newWorkbook.Sheets[1];

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
                        newSheet.Cells[1, 1],
                        newSheet.Cells[endRow - counter + 1, colCount]
                        ];

                    sourceRange.Copy(destRange);

                    // Save new workbook
                    string newFile = Path.Combine(outputFolder, $"Batch_{fileIndex}.xlsx");
                    newWorkbook.SaveAs(newFile);
                    newWorkbook.Close();

                    Console.WriteLine($"Saved {newFile}");

                    // Release batch ranges inside each iteration of while loop
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
                //sourceWorkbook?.Close(false);
                //excelApp.Quit();

                //System.Runtime.InteropServices.Marshal.ReleaseComObject(sourceSheet);
                //System.Runtime.InteropServices.Marshal.ReleaseComObject(sourceWorkbook);
                //System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

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
