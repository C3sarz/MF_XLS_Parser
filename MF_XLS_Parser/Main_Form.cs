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
using Microsoft.Office.Interop.Excel;

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

        /// <summary>
        /// Range worked on in the excel file.
        /// </summary>
        private Excel.Range xlRange;

        /// <summary>
        /// Instance of the created output Excel app.
        /// </summary>
        Excel.Application newExcelApp;

        /// <summary>
        /// New workbook created to copy data into (output).
        /// </summary>
        Excel._Workbook newWorkBook;

        /// <summary>
        /// The new sheet created in the output workbook.
        /// </summary>
        Excel._Worksheet newSheet;

        /// <summary>
        /// First background worker.
        /// </summary>
        private BackgroundWorker backgroundWorker1 = new BackgroundWorker();

        /// <summary>
        /// Second background worker.
        /// </summary>
        private BackgroundWorker backgroundWorker2 = new BackgroundWorker();

        /// <summary>
        /// Third background worker.
        /// </summary>
        private BackgroundWorker backgroundWorker3 = new BackgroundWorker();

        /// <summary>
        /// Tracks the amount of completed working threads.
        /// </summary>
        private int workersCompleted;

        private bool columnsReady = false;
        private int[] columns;
        private int firstDataRow;
        private string fileName;

        /// <summary>
        /// Form constructor.
        /// </summary>
        public Main_Form()
        {
            InitializeComponent();

            //Worker setup.
            backgroundWorker1.DoWork += BackgroundWorker1_DoWork;
            backgroundWorker2.DoWork += BackgroundWorker2_DoWork;
            backgroundWorker3.DoWork += BackgroundWorker3_DoWork;
            backgroundWorker1.RunWorkerCompleted += BackgroundWorkers_RunWorkerCompleted;
            backgroundWorker2.RunWorkerCompleted += BackgroundWorkers_RunWorkerCompleted;
            backgroundWorker3.RunWorkerCompleted += BackgroundWorkers_RunWorkerCompleted;
        }

        /// <summary>
        /// "Open File" button event handler.
        /// </summary>
        /// <param name="sender">The button.</param>
        /// <param name="e">Event.</param>
        private void OpenFileButton_Click(object sender, EventArgs e)
        {
            //File input.
            OpenFileDialog openDialog = new OpenFileDialog();
            if (openDialog.ShowDialog() == DialogResult.OK)
            {
                OpenFileButton.Enabled = false;
                fileName = openDialog.FileName;
                DataTextBox.Text = fileName;
                AppLoadingImage.Visible = true;
                columnsReady = false;
                backgroundWorker1.RunWorkerAsync();
            }
        }

        /// <summary>
        /// "Show Cell Contents" button event handler.
        /// </summary>
        /// <param name="sender">The button.</param>
        /// <param name="e">Event.</param>
        private void ShowCellButton_Click(object sender, EventArgs e)
        {
            int i = Int32.Parse(cellBox1.Text);
            int j = Int32.Parse(cellBox2.Text);
            try
            {
                if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                {
                    DataTextBox.Text = (string)(xlRange.Cells[i, j] as Excel.Range).Value2;
                }
                else DataTextBox.Text = "Celda Vacia";
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
            StartButton.Enabled = false;
            try
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
                DataTextBox.Clear();
            }

            catch (Exception ex)
            {
                MessageBox.Show("Error: No hay documento para cerrar.");
            }
        }

        /// <summary>
        /// Write Cell button event handler.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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
        private void StartButton_Click(object sender, EventArgs e)
        {
            if (columnsReady)
            {
                //Workbook
                StartButton.Enabled = false;
                LoadingImage.Visible = true;
                newExcelApp = new Excel.Application();
                newExcelApp.Visible = true;
                newWorkBook = (Excel._Workbook)(newExcelApp.Workbooks.Add(Missing.Value));
                newSheet = (Excel._Worksheet)newWorkBook.ActiveSheet;
                //Sheet setup
                newSheet.Cells[1, 1] = "Codigo";
                newSheet.Cells[1, 2] = "Producto";
                newSheet.Cells[1, 3] = "Cantidad";
                newSheet.Cells[1, 4] = "Total";

                //Launch worker threads.
                workersCompleted = 0;
                backgroundWorker1.RunWorkerAsync();
                backgroundWorker2.RunWorkerAsync();
            }
            else MessageBox.Show("Por favor confimar la primera fila de datos.");
        }

        private void FilterStartButton_Click(object sender, EventArgs e)
        {
            if (columnsReady && FilterBox.Text != "")
            {
                //Workbook
                StartButton.Enabled = false;
                FilterStartButton.Enabled = false;
                LoadingImage.Visible = true;
                newExcelApp = new Excel.Application();
                newExcelApp.Visible = true;
                newWorkBook = (Excel._Workbook)(newExcelApp.Workbooks.Add(Missing.Value));
                newSheet = (Excel._Worksheet)newWorkBook.ActiveSheet;
                //Sheet setup
                newSheet.Cells[1, 1] = "Codigo";
                newSheet.Cells[1, 2] = "Producto";
                newSheet.Cells[1, 3] = "Cantidad";
                newSheet.Cells[1, 4] = "Total";

                //Launch worker threads.
                workersCompleted = 1;
                backgroundWorker3.RunWorkerAsync();
            }
            else MessageBox.Show("Por favor confimar fila y filtro");
        }

        /// <summary>
        /// Transfers a specified column on the Excel file into a new one, skipping null cells.
        /// </summary>
        /// <param name="startingRow">Row to start parsing.</param>
        /// <param name="parsedColumn">Column to be parsed.</param>
        /// <param name="newSheetPositionX">Starting column cell in which the parsed data is copied.</param>
        /// <param name="newSheetPositionY">Starting row cell in which the parsed data is copied.</param>
        private void startTransfer(int startingRow, int parsedColumn, int newSheetPositionX, int newSheetPositionY)
        {
            try
            {
                int nullCount = 0;
                int currentPosition = startingRow;

                // Iteration through cells.
                while (nullCount < 10)
                {
                    if (xlRange.Cells[currentPosition, parsedColumn] == null || xlRange.Cells[currentPosition, parsedColumn].Value2 == null)
                    {
                        nullCount++;
                    }
                    else
                    {
                        nullCount = 0;
                        newSheet.Cells[newSheetPositionY, newSheetPositionX] = (xlRange.Cells[currentPosition, parsedColumn]).Value2;
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
        /// Transfers a specified column on the Excel file into a new one, skipping null cells.
        /// </summary>
        /// <param name="startingRow">Row to start parsing.</param>
        /// <param name="newSheetPositionX">Starting column cell in which the parsed data is copied.</param>
        /// <param name="newSheetPositionY">Starting row cell in which the parsed data is copied.</param>
        private void startFullTransfer(int startingRow, int newSheetPositionX, int newSheetPositionY)
        {
            try
            {
                int nullCount = 0;
                int currentPosition = startingRow;
                int namesColumn = columns[2];
                int quantityColumn = columns[3];

                // Iteration through cells.
                while (nullCount < 10)
                {
                    if (xlRange.Cells[currentPosition, namesColumn] == null || xlRange.Cells[currentPosition, namesColumn].Value2 == null)
                    {
                        nullCount++;
                    }
                    else
                    {
                        nullCount = 0;
                        //Name copying.
                        newSheet.Cells[newSheetPositionY, newSheetPositionX] = (xlRange.Cells[currentPosition, namesColumn]).Value2;

                        //Quantity copying.
                        newSheet.Cells[newSheetPositionY, newSheetPositionX + 1] = (xlRange.Cells[currentPosition, quantityColumn]).Value2.ToString();

                        //Total copying.
                        newSheet.Cells[newSheetPositionY, newSheetPositionX + 2] = (xlRange.Cells[currentPosition, quantityColumn + 2]).Value2.ToString();

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
        /// Transfers a specified column on the Excel file into a new one, skipping null cells and filtering unwanted items.
        /// </summary>
        /// <param name="startingRow">Row to start parsing.</param>
        /// <param name="newSheetPositionX">Starting column cell in which the parsed data is copied.</param>
        /// <param name="newSheetPositionY">Starting row cell in which the parsed data is copied.</param>
        /// <param name="filterTerm">Filter term</param>
        private void startFullTransfer(int startingRow, int newSheetPositionX, int newSheetPositionY, string filterTerm)
        {
            try
            {
                int nullCount = 0;
                int currentPosition = startingRow;
                int namesColumn = columns[2];
                int quantityColumn = columns[3];

                // Iteration through cells.
                while (nullCount < 10)
                {
                    if (xlRange.Cells[currentPosition, namesColumn] == null || xlRange.Cells[currentPosition, namesColumn].Value2 == null)
                    {
                        nullCount++;
                    }
                    else
                    {
                        nullCount = 0;
                        string currentString = xlRange.Cells[currentPosition, namesColumn].Value2;
                        if (currentString.Contains(filterTerm))
                        {
                            //Name copying.
                            newSheet.Cells[newSheetPositionY, newSheetPositionX] = (xlRange.Cells[currentPosition, namesColumn]).Value2;

                            //Quantity copying.
                            newSheet.Cells[newSheetPositionY, newSheetPositionX + 1] = (xlRange.Cells[currentPosition, quantityColumn]).Value2.ToString();

                            //Total copying.
                            newSheet.Cells[newSheetPositionY, newSheetPositionX + 2] = (xlRange.Cells[currentPosition, quantityColumn + 2]).Value2.ToString();


                            newSheetPositionY++;
                        }
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

        private int[] getDataColumns(int firstDataRow)
        {
            int processedColumn = 0;
            int position = 1;
            int[] dataColumns = new int[5];
            while (processedColumn < 5)
            {
                if ((xlRange.Cells[firstDataRow, position]).Value2 != null)
                {
                    if (processedColumn == 0 && !Double.TryParse((xlRange.Cells[firstDataRow, position].Value2.ToString()), out double result))
                    {
                        throw new Exception("Could not find columns");
                    }
                    dataColumns[processedColumn] = position;
                    processedColumn++;
                }
                position++;
                if (position >= 100) throw new Exception("No se encontraron columnas de datos.");
            }
            this.firstDataRow = firstDataRow;
            return dataColumns;
        }




        /// <summary>
        /// Background worker 1 work method.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BackgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            if (columnsReady)
            {
                int newSheetPositionX = 1;
                int newSheetPositionY = 3;
                int parsedColumn = columns[0];
                //Transfer codes
                startTransfer(firstDataRow, parsedColumn, newSheetPositionX, newSheetPositionY);
            }
            else
            {

                excelApp = new Excel.Application();
                excelApp.Visible = false;
                currentWorkbook = excelApp.Workbooks.Open(@fileName);
                currentSheet = (Excel.Worksheet)currentWorkbook.Worksheets.get_Item(1);
                xlRange = currentSheet.UsedRange;

            }
        }

        /// <summary>
        /// Background worker 2 work method.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BackgroundWorker2_DoWork(object sender, DoWorkEventArgs e)
        {
            int newSheetPositionX = 2;
            int newSheetPositionY = 3;
            startFullTransfer(firstDataRow, newSheetPositionX, newSheetPositionY);
        }

        /// <summary>
        /// Background worker 3 work method.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BackgroundWorker3_DoWork(object sender, DoWorkEventArgs e)
        {
            int newSheetPositionX = 2;
            int newSheetPositionY = 3;


            try
            {
                startFullTransfer(firstDataRow, newSheetPositionX, newSheetPositionY, FilterBox.Text);
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
            if (columnsReady)
            {
                workersCompleted++;
                if (workersCompleted >= 2)
                {
                    StartButton.Enabled = true;
                    LoadingImage.Visible = false;
                    FilterStartButton.Enabled = true;
                    columnsReady = false;
                    RowBox.BackColor = Color.White;
                    MessageBox.Show("Proceso completado");
                }
            }
            else
            {
                AppLoadingImage.Visible = false;
                StartButton.Enabled = true;
                FilterStartButton.Enabled = true;
                OpenFileButton.Enabled = true;
            }
        }

        private void RowConfirmButton_Click(object sender, EventArgs e)
        {
            try
            {
                columns = getDataColumns(Int32.Parse(RowBox.Text));
                columnsReady = true;
                RowBox.BackColor = Color.LightGreen;
            }
            catch (Exception ex)
            {
                columnsReady = false;
                RowBox.BackColor = Color.White;
                MessageBox.Show("Error confirmando fila.");
            }
        }


    }
}
