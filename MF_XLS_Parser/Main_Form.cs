using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Reflection;
using System.IO;
using System.Security.Authentication;
using Microsoft.CSharp.RuntimeBinder;

namespace MF_XLS_Parser
{
    public partial class Main_Form : Form
    {
        /// <summary>
        /// Instance of the Excel application.
        /// </summary>
        private Excel.Application excelApp;

        /// <summary>
        /// Workbook currently worked on.
        /// </summary>
        private Excel._Workbook currentWorkbook;

        /// <summary>
        /// Current sheet on the workbook.
        /// </summary>
        private Excel._Worksheet currentSheet;
        private Excel.Range xlRange;
        private int workersCompleted = 0;

        Excel.Application newExcelApp;
        Excel._Workbook newWorkBook;
        Excel._Worksheet newSheet;

        BackgroundWorker backgroundWorker1 = new BackgroundWorker();
        BackgroundWorker backgroundWorker2 = new BackgroundWorker();
        public Main_Form()
        {
            InitializeComponent();

            backgroundWorker1.DoWork += BackgroundWorker1_DoWork;
            backgroundWorker2.DoWork += BackgroundWorker2_DoWork;
            backgroundWorker1.RunWorkerCompleted += BackgroundWorkers_RunWorkerCompleted;
            backgroundWorker2.RunWorkerCompleted += BackgroundWorkers_RunWorkerCompleted;
        }

        /// <summary>
        /// "Open File" button event handler.
        /// </summary>
        /// <param name="sender">The button.</param>
        /// <param name="e">Event.</param>
        private void Button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openDialog = new OpenFileDialog();

            if (openDialog.ShowDialog() == DialogResult.OK)
            {
                string fileName = openDialog.FileName;
                DataTextBox.Text = fileName;
                excelApp = new Excel.Application();
                excelApp.Visible = false;
                currentWorkbook = excelApp.Workbooks.Open(@fileName);
                currentSheet = (Excel.Worksheet)currentWorkbook.Worksheets.get_Item(1);
                xlRange = currentSheet.UsedRange;
            }
        }

        /// <summary>
        /// "Show Cell Contents" button event handler.
        /// </summary>
        /// <param name="sender">The button.</param>
        /// <param name="e">Event.</param>
        private void button2_Click(object sender, EventArgs e)
        {
            int i = Int32.Parse(cellBox1.Text);
            int j = Int32.Parse(cellBox2.Text);
            try
            {
                if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                {
                    DataTextBox.Text = (string)(xlRange.Cells[i, j] as Excel.Range).Value2;
                }
                else DataTextBox.Text = "Empty Cell";
            }

            catch (RuntimeBinderException ex)
            {
                DataTextBox.Text = ((double)(xlRange.Cells[i, j] as Excel.Range).Value2).ToString();
            }

            catch (Exception ex)
            {
                String errorMessage;
                errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, ex.Message);
                errorMessage = String.Concat(errorMessage, "\n Full String: ");
                errorMessage = String.Concat(errorMessage, ex.ToString());
                MessageBox.Show(errorMessage, "Error");
            }
            //if (xlRange.Cells[i, j].Value2 == null) InfoTextBox.Text = "null";
            //else InfoTextBox.Text = (xlRange.Cells[i, j] as Excel.Range).Value2.GetType().ToString();
        }

        /// <summary>
        /// Cleanup button event handler.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CleanupButton_Click(object sender, EventArgs e)
        {
            Cleanup();
        }

        /// <summary>
        /// Gets rid of background threads.
        /// </summary>
        private void Cleanup()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //Original file cleanup
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(currentSheet);
            currentWorkbook.Close();
            Marshal.ReleaseComObject(currentWorkbook);
            excelApp.Quit();
            Marshal.ReleaseComObject(excelApp);
        }

        private void WriteButton_Click(object sender, EventArgs e)
        {
            int i = Int32.Parse(cellBox1.Text);
            int j = Int32.Parse(cellBox2.Text);
            (xlRange.Cells[i, j] as Excel.Range).Value2 = DataTextBox.Text;
        }

        /// <summary>
        /// Start removal of unecessary cells and parse old cells to a new file.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ParsingButton_Click(object sender, EventArgs e)
        {
            ParsingButton.Enabled = false;
            LoadingImage.Visible = true;
            newExcelApp = new Excel.Application();
            newExcelApp.Visible = true;
            newWorkBook = (Excel._Workbook)(newExcelApp.Workbooks.Add(Missing.Value));
            newSheet = (Excel._Worksheet)newWorkBook.ActiveSheet;

            backgroundWorker1.RunWorkerAsync();
            backgroundWorker2.RunWorkerAsync();
        }

        /// <summary>
        /// Background worker 1 work method.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BackgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                int nullCount = 0;
                int currentPosition = 9;
                int newSheetPositionX = 1;
                int newSheetPositionY = 1;

                // Iteration through cells.
                while (nullCount < 10)
                {
                    if (xlRange.Cells[currentPosition, 2] == null || xlRange.Cells[currentPosition, 2].Value2 == null)
                    {
                        nullCount++;
                    }
                    else
                    {
                        nullCount = 0;
                        newSheet.Cells[newSheetPositionY, newSheetPositionX] = (string)(xlRange.Cells[currentPosition, 2] as Excel.Range).Value2;
                        newSheetPositionY++;
                    }
                    currentPosition++;
                }
            }

            catch (Exception ex)
            {
                String errorMessage;
                errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, ex.Message);
                errorMessage = String.Concat(errorMessage, "\n Full String: ");
                errorMessage = String.Concat(errorMessage, ex.ToString());
                MessageBox.Show(errorMessage, "Error");
            }
        }

        /// <summary>
        /// Background worker 2 work method.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BackgroundWorker2_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                int nullCount = 0;
                int currentPosition = 9;
                int newSheetPositionX = 3;
                int newSheetPositionY = 1;

                // Iteration through cells.
                while (nullCount < 10)
                {
                    if (xlRange.Cells[currentPosition, 19] == null || xlRange.Cells[currentPosition, 19].Value2 == null)
                    {
                        nullCount++;
                    }
                    else
                    {
                        nullCount = 0;
                        newSheet.Cells[newSheetPositionY, newSheetPositionX] = (string)(xlRange.Cells[currentPosition, 19] as Excel.Range).Value2;
                        newSheetPositionY++;
                    }
                    currentPosition++;
                }
            }

            //Error handling
            catch (Exception ex)
            {
                String errorMessage;
                errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, ex.Message);
                errorMessage = String.Concat(errorMessage, "\n Full String: ");
                errorMessage = String.Concat(errorMessage, ex.ToString());
                MessageBox.Show(errorMessage, "Error");
            }
        }

        /// <summary>
        /// Work completed event handler
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BackgroundWorkers_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            workersCompleted++;
            if (workersCompleted >= 2)
            {
                ParsingButton.Enabled = true;
                LoadingImage.Visible = false;
                MessageBox.Show("Parsing completado");
            }
        }
    }
}
