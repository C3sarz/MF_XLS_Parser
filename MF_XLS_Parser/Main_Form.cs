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
using System.Collections.Concurrent;

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

        private BackgroundWorker backgroundWorker4 = new BackgroundWorker();

        /// <summary>
        /// Tracks the amount of completed working threads.
        /// </summary>
        private int workersCompleted;

        private bool dataColumnsReady = false;
        private bool typeColumnsReady = false;
        public bool threadsWorking = false;
        private int[] dataColumns;
        private int[] typeColumns;
        private int[] startingRows;
        private string fileName;
        private State state = State.Idle;
        private ConcurrentQueue<NamesBlock> namesQ = new ConcurrentQueue<NamesBlock>();
        private ConcurrentQueue<DataBlock> numbersQ = new ConcurrentQueue<DataBlock>();

        public enum State
        {
            Idle,
            Testing,
            CodeProcessing,
            FullProcessing,
            LoadingFile
        }

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
            backgroundWorker4.DoWork += BackgroundWorker4_DoWork;
            backgroundWorker1.RunWorkerCompleted += BackgroundWorkers_RunWorkerCompleted;
            backgroundWorker2.RunWorkerCompleted += BackgroundWorkers_RunWorkerCompleted;
            backgroundWorker3.RunWorkerCompleted += BackgroundWorkers_RunWorkerCompleted;
            backgroundWorker4.RunWorkerCompleted += BackgroundWorkers_RunWorkerCompleted;

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
                state = State.LoadingFile;
                OpenFileButton.Enabled = false;
                fileName = openDialog.FileName;
                DataTextBox.Text = fileName;
                AppLoadingImage.Visible = true;
                dataColumnsReady = false;
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
            if (dataColumnsReady)
            {
                state = State.FullProcessing;
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
                newSheet.Cells[1, 5] = "Seccion";
                newSheet.Cells[1, 6] = "Grupo";
                newSheet.Cells[1, 7] = "Categoria";
                newSheet.Cells[1, 8] = "Sub-Categoria";

                //Launch worker threads.
                workersCompleted = 0;
                backgroundWorker1.RunWorkerAsync();
                backgroundWorker2.RunWorkerAsync();
            }
            else MessageBox.Show("Por favor confimar la primera fila de datos.");
        }

        private void FilterStartButton_Click(object sender, EventArgs e)
        {
            if (dataColumnsReady && FilterBox.Text != "")
            {
                state = State.FullProcessing;
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

        private void TestButton_Click(object sender, EventArgs e)
        {
            if (dataColumnsReady)
            {
                state = State.Testing;
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
                (newSheet.Cells[1, 2] as Excel.Range).ColumnWidth = 45;
                newSheet.Cells[1, 3] = "Cantidad";
                newSheet.Cells[1, 4] = "Total";
                newSheet.Cells[1, 5] = "Seccion";
                newSheet.Cells[1, 6] = "Grupo";
                newSheet.Cells[1, 7] = "Categoria";
                newSheet.Cells[1, 8] = "Sub-Categoria   ";

                //Launch worker threads.
                workersCompleted = 0;
                threadsWorking = true;
                backgroundWorker2.RunWorkerAsync();
                backgroundWorker3.RunWorkerAsync();
            }
            else MessageBox.Show("Por favor confimar la primera fila de datos.");
        }

        private void startNamesCopy(NamesBlock names)
        {
            try
            {
                // Iteration through cells.
                for(int i = names.StartRow; i <= names.EndRow;i++)
                {
                    newSheet.Cells[i, 5] = names.Section;
                    newSheet.Cells[i, 6] = names.Group;
                    newSheet.Cells[i, 7] = names.Category;
                    newSheet.Cells[i, 8] = names.SubCategory;

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
        private void startFullTransfer(int newSheetPositionX, int newSheetPositionY)
        {
            try
            {
                int nullCount = 0;
                bool dataCopied = false;
                int startY = newSheetPositionY;

                int currentPosition = startingRows[0];
                int namesColumn = dataColumns[2];
                int quantityColumn = dataColumns[3];
                int sectionColumn = typeColumns[0];
                int groupColumn = typeColumns[0] + 1;
                int categoryColumn = typeColumns[0] + 2;
                int subCategoryColumn = typeColumns[0] + 3;
                string section = xlRange.Cells[startingRows[1], typeColumns[1]].Value2;
                string group = xlRange.Cells[startingRows[1] + 1, typeColumns[1] + 2].Value2;
                string category = xlRange.Cells[startingRows[1] + 2, typeColumns[1] + 4].Value2;
                string subCategory = xlRange.Cells[startingRows[1] + 3, typeColumns[1] + 6].Value2;


                // Iteration through cells.
                while (nullCount < 10)
                {
                    if (xlRange.Cells[currentPosition, namesColumn] == null || xlRange.Cells[currentPosition, namesColumn].Value2 == null)
                    {
                        if (!dataCopied)
                        {
                            //Type copying
                            namesQ.Enqueue(new NamesBlock(startY, newSheetPositionY - 1, section, group, category, subCategory));

                            dataCopied = true;
                        }

                        int type = findUsedColumn(currentPosition, 1);
                        if (type == sectionColumn)
                        {
                            section = (xlRange.Cells[currentPosition, sectionColumn + 8]).Value2;
                        }
                        else if (type == groupColumn)
                        {
                            group = (xlRange.Cells[currentPosition, groupColumn + 9]).Value2;
                        }
                        else if (type == categoryColumn)
                        {
                            category = (xlRange.Cells[currentPosition, categoryColumn + 10]).Value2;
                        }
                        else if (type == subCategoryColumn)
                        {
                            subCategory = (xlRange.Cells[currentPosition, subCategoryColumn + 11]).Value2;
                        }

                        nullCount++;
                    }
                    else
                    {
                        if (dataCopied) startY = newSheetPositionY;
                        dataCopied = false;
                        nullCount = 0;
                        

                        //Name copying.
                        newSheet.Cells[newSheetPositionY, newSheetPositionX] = (xlRange.Cells[currentPosition, namesColumn]).Value2;

                        //Code copying.
                        newSheet.Cells[newSheetPositionY, newSheetPositionX - 1] = (xlRange.Cells[currentPosition, dataColumns[0]]).Value2;

                        //Quantity copying.
                        newSheet.Cells[newSheetPositionY, newSheetPositionX + 1] = (xlRange.Cells[currentPosition, quantityColumn]).Value2.ToString();

                        //Total copying.
                        newSheet.Cells[newSheetPositionY, newSheetPositionX + 2] = (xlRange.Cells[currentPosition, quantityColumn + 2]).Value2.ToString();

                        //Type copying
                        //newSheet.Cells[newSheetPositionY, newSheetPositionX + 3] = section;
                        //newSheet.Cells[newSheetPositionY, newSheetPositionX + 4] = group;
                        //newSheet.Cells[newSheetPositionY, newSheetPositionX + 5] = category;
                        //newSheet.Cells[newSheetPositionY, newSheetPositionX + 6] = subCategory;


                        newSheetPositionY++;
                        //if (state == State.Testing && newSheetPositionY > 100) break;
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
        private void startFullTransfer(int newSheetPositionX, int newSheetPositionY, string filterTerm)
        {
            try
            {
                int nullCount = 0;
                int currentPosition = startingRows[0];
                int namesColumn = dataColumns[2];
                int quantityColumn = dataColumns[3];

                // Iteration through cells.
                while (nullCount < 10)
                {
                    //if (xlRange.Cells[currentPosition, namesColumn] == null || xlRange.Cells[currentPosition, namesColumn].Value2 == null)
                    if (xlRange.Cells[currentPosition, namesColumn] == null)
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

        private void getColumns(int firstDataRow, int firstTypeRow)
        {
            int processedColumn = 0;
            int position = 1;
            dataColumns = new int[5];
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
            startingRows = new int[2];
            startingRows[0] = firstDataRow;
            startingRows[1] = firstTypeRow;

            processedColumn = 0;
            position = 1;
            typeColumns = new int[4];
            while (processedColumn < 4)
            {
                if ((xlRange.Cells[firstTypeRow, position]).Value2 != null)
                {
                    typeColumns[processedColumn] = position;
                    processedColumn++;
                }
                position++;
                if (position >= 100) throw new Exception("No se encontraron columnas de datos.");
            }
        }

        /// <summary>
        /// Background worker 1 work method.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BackgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            if (dataColumnsReady)
            {

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
            if (state == State.FullProcessing)
            {
                int newSheetPositionX = 2;
                int newSheetPositionY = 3;
                startFullTransfer(newSheetPositionX, newSheetPositionY);
            }
            else if (state == State.Testing)
            {
                int newSheetPositionX = 2;
                int newSheetPositionY = 3;
                startFullTransfer(newSheetPositionX, newSheetPositionY);
            }
            threadsWorking = false;
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
            NamesBlock names;

            try
            {
                while(threadsWorking)
                {
                    if(namesQ.TryDequeue(out names))
                    {
                        startNamesCopy(names);
                    }
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
        /// Background worker 4 work method.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BackgroundWorker4_DoWork(object sender, DoWorkEventArgs e)
        {
            int newSheetPositionX = 2;
            int newSheetPositionY = 3;


            try
            {

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
            if (dataColumnsReady)
            {
                workersCompleted++;
                if (workersCompleted >= 2)
                {
                    StartButton.Enabled = true;
                    LoadingImage.Visible = false;
                    FilterStartButton.Enabled = true;
                    dataColumnsReady = false;
                    typeColumnsReady = false;
                    RowBox1.BackColor = Color.White;
                    RowBox2.BackColor = Color.White;
                    //Cleanup();
                    MessageBox.Show("Proceso completado");
                }
            }
            else
            {
                AppLoadingImage.Visible = false;
                StartButton.Enabled = true;
                TestButton.Enabled = true;
                FilterStartButton.Enabled = true;
                OpenFileButton.Enabled = true;
            }
            state = State.Idle;
        }

        /// <summary>
        /// Confirms the row in order to start the processing.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void RowConfirmButton_Click(object sender, EventArgs e)
        {
            try
            {
                getColumns(Int32.Parse(RowBox1.Text), Int32.Parse(RowBox2.Text));
                dataColumnsReady = true;
                RowBox1.BackColor = Color.LightGreen;
                typeColumnsReady = true;
                RowBox2.BackColor = Color.LightGreen;
            }
            catch (Exception ex)
            {
                dataColumnsReady = false;
                RowBox1.BackColor = Color.White;
                typeColumnsReady = false;
                RowBox2.BackColor = Color.White;
                MessageBox.Show("Error confirmando filas.");
            }
        }

        /// <summary>
        /// Finds the next non-null column in a row.
        /// </summary>
        /// <param name="row">Row to be searched.</param>
        /// <param name="startCol">Starting column index.</param>
        /// <returns></returns>
        public int findUsedColumn(int row, int startCol)
        {
            int currentCol = startCol;
            while (xlRange.Cells[row, currentCol] == null || xlRange.Cells[row, currentCol].Value2 == null)
            {
                currentCol++;
                if (currentCol >= 70) return -1;
            }
            return currentCol;
        }
    }
}
