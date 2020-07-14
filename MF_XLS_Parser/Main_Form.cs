/*
 * XLS Parser MF
 * Main_Form.cs
 * Author: Cesar Zavala
 * 
 */
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
using System.Windows.Forms.VisualStyles;
using System.Diagnostics;
using System.CodeDom;

namespace MF_XLS_Parser
{
    public partial class Main_Form : Form
    {
        ExcelData input;
        ExcelData output;

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
        /// Enables the lopp worker thread.
        /// </summary>
        public bool workerThreadEnabled = false;

        /// <summary>
        /// Rows where the data starts.
        /// </summary>
        private int[] startingRows;

        /// <summary>
        /// Keeps track of the state of the program.
        /// </summary>
        private State state = State.Idle;

        /// <summary>
        /// Keeps track of all active threads.
        /// </summary>
        int activeThreads;

        /// <summary>
        /// Queue that keeps track of data types for multithreading.
        /// </summary>
        private ConcurrentQueue<NamesBlock> namesQ = new ConcurrentQueue<NamesBlock>();

        /// <summary>
        /// Determines if a file is loaded.
        /// </summary>
        public bool IsFileLoaded { get; private set; } = false;

        /// <summary>
        /// Stores the input file location.
        /// </summary>
        private string fileName;

        /// <summary>
        /// List set of the product names.
        /// </summary>
        public HashSet<string> NamesList { get; private set; }

        public Dictionary<string, string> FormulaStrings;

        /// <summary>
        /// Keeps track of the elapsed time.
        /// </summary>
        private Stopwatch timer = new Stopwatch();

        /// <summary>
        /// Duplicates found while processing an excel file.
        /// </summary>
        int duplicatesProcessed = 0;

        /// <summary>
        /// Duplicates found in input file
        /// </summary>
        int inputDuplicates = 0;

        /// <summary>
        /// States of the program.
        /// </summary>
        public enum State
        {
            Idle,
            Testing,
            FilterProcessing,
            FullProcessing,
            LoadingFile
        }

        /// <summary>
        /// Form constructor.
        /// </summary>
        public Main_Form()
        {
            InitializeComponent();

            //Event Handlers
            this.FormClosing += new FormClosingEventHandler(Closing);

            //Worker setup.
            backgroundWorker1.DoWork += BackgroundWorker1_DoWork;
            backgroundWorker2.DoWork += BackgroundWorker2_DoWork;
            backgroundWorker3.DoWork += BackgroundWorker3_DoWork;
            backgroundWorker1.RunWorkerCompleted += BackgroundWorkers_RunWorkerCompleted;
            backgroundWorker2.RunWorkerCompleted += BackgroundWorkers_RunWorkerCompleted;
            backgroundWorker3.RunWorkerCompleted += BackgroundWorkers_RunWorkerCompleted;
            backgroundWorker2.ProgressChanged += Worker2_ProgressChanged;
            backgroundWorker2.WorkerReportsProgress = true;

        }

        /// <summary>
        /// "Open File" button event handler.
        /// </summary>
        /// <param name="sender">The button.</param>
        /// <param name="e">Event.</param>
        private void OpenFileButton_Click(object sender, EventArgs e)
        {
            //File input.
            try
            {
                OpenFileDialog openDialog = new OpenFileDialog();
                if (openDialog.ShowDialog() == DialogResult.OK)
                {
                    if (IsFileLoaded) Cleanup();
                    state = State.LoadingFile;
                    fileName = openDialog.FileName;
                    OpenFileButton.Enabled = false;
                    DataTextBox.Text = "Archivo cargado: \n\n" + openDialog.FileName;
                    AppLoadingImage.Visible = true;
                    backgroundWorker1.RunWorkerAsync();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error cargando el archivo, verifique el formato.");
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
                if (input.fullRange.Cells[i, j] != null && input.fullRange.Cells[i, j].Value2 != null)
                {
                    DataTextBox.Text = (string)(input.fullRange.Cells[i, j] as Excel.Range).Value2;
                }
                else DataTextBox.Text = "Celda Vacia";
            }

            catch (RuntimeBinderException ex)
            {
                DataTextBox.Text = ((double)(input.fullRange.Cells[i, j] as Excel.Range).Value2).ToString();
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
            TestButton.Enabled = false;
            FilterStartButton.Enabled = false;
            try
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();

                //Original file cleanup
                Marshal.ReleaseComObject(input.fullRange);
                Marshal.ReleaseComObject(input.currentSheet);
                input.currentWorkbook.Close();
                Marshal.ReleaseComObject(input.currentWorkbook);
                DataTextBox.Clear();
                ListBox.Clear();
                IsFileLoaded = false;
            }

            catch (Exception ex)
            {
                MessageBox.Show("Error: No hay documento para cerrar.");
            }
        }

        /// <summary>
        /// Start removal of unecessary cells and parse old cells to a new file.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void StartButton_Click(object sender, EventArgs e)
        {
            if (input.dataColumnsReady)
            {
                state = State.FullProcessing;
                //Workbook
                StartButton.Enabled = false;
                LoadingImage.Visible = true;
                output = new ExcelData(null);

                //Launch worker threads.
                activeThreads = 2;
                workerThreadEnabled = true;
                backgroundWorker2.RunWorkerAsync();
                backgroundWorker3.RunWorkerAsync();
            }
            else MessageBox.Show("Por favor confimar la primera fila de datos.");
        }

        private void FilterStartButton_Click(object sender, EventArgs e)
        {
            if (input.dataColumnsReady
                && input.typeColumnsReady)
            {
                state = State.FilterProcessing;
                //Workbook
                StartButton.Enabled = false;
                FilterStartButton.Enabled = false;
                LoadingImage.Visible = true;
                output = new ExcelData(null);

                try
                {
                    //Read file from dialog.
                    NamesList = new HashSet<string>();
                    FormulaStrings = new Dictionary<string, string>();
                    OpenFileDialog namesDialog = new OpenFileDialog();
                    OpenFileDialog formulaDialog = new OpenFileDialog();
                    if (namesDialog.ShowDialog() == DialogResult.OK)
                    {
                        if (formulaDialog.ShowDialog() == DialogResult.OK)
                        {
                            StreamReader sr1 = new StreamReader(namesDialog.FileName);

                            StreamReader sr2 = new StreamReader(formulaDialog.FileName);
                            string lineA;
                            string lineB;
                            int count = 1;
                            while ((lineA = sr1.ReadLine()) != null)
                            {
                                lineA = lineA.ToUpper();
                                lineB = sr2.ReadLine();
                                if (NamesList.Contains(lineA))
                                {
                                    //do nothing
                                    inputDuplicates++;
                                }
                                else
                                {
                                    NamesList.Add(lineA);
                                    FormulaStrings.Add(lineA, lineB);

                                }
                                count++;
                            }
                            sr1.Close();
                            sr2.Close();
                        }
                        else throw new Exception("Error loading file.");
                    }
                    else throw new Exception("Error loading file.");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error cargando el archivo, verifique el formato.");
                }

                //Launch worker threads.
                activeThreads = 1;
                workerThreadEnabled = false;
                backgroundWorker2.RunWorkerAsync();
            }
            else MessageBox.Show("Por favor confimar fila y filtro");
        }

        private double getUnitaryPrice(string name, double quantity, double total)
        {
            double value = total / quantity;
            string formula;
            if (FormulaStrings.TryGetValue(name, out formula))
            {
                //If there is no extra operation.
                if (formula.Length == 23)
                {
                    return value;
                }

                //If the formula has '*'
                else if (formula.Contains('*'))
                {
                    //Split into parts
                    string[] parts = formula.Split(new char[] { '*' });

                    //If there are no operations after parenthesis
                    if (parts[1].Length - 1 == parts[1].LastIndexOf(')'))
                    {
                        parts[1] = parts[1].Replace("(", "");
                        parts[1] = parts[1].Replace(")", "");
                        string[] subparts = parts[1].Split('/');
                        double a = Double.Parse(subparts[0]);
                        double b = Double.Parse(subparts[1]);

                        return value * (a / b);
                    }
                    //If there is division after parenthesis.
                    else
                    {
                        string[] subparts = parts[1].Split(new char[] { '/', '(', ')' });
                        double a = Double.Parse(subparts[1]);
                        double b = Double.Parse(subparts[2]);
                        //Replace commas by dots for parsing.
                        subparts[4] = subparts[4].Replace(",", "."); 
                        double c = Double.Parse(subparts[4]);
                        return value * (a / b) / c;
                    }
                }
                //If only division after base value.
                else if (formula[23] == '/')
                {
                    string[] parts = formula.Split(new char[] { '/', '(', ')' });
                    //Replace commas by dots for parsing.
                    parts[2] = parts[2].Replace(",", ".");
                    double a = Double.Parse(parts[2]);
                    return value / a;
                }
                else
                {
                    throw new Exception("Formula parsing failed.");
                }
            }
            else
            {
                throw new Exception("Formula not found.");
            }
        }

        private void TestButton_Click(object sender, EventArgs e)
        {
            if (input.dataColumnsReady && input.typeColumnsReady)
            {
                state = State.Testing;
                //Workbook
                StartButton.Enabled = false;
                LoadingImage.Visible = true;
                output = new ExcelData(null);

                //Launch worker threads.
                activeThreads = 2;
                workerThreadEnabled = true;
                backgroundWorker2.RunWorkerAsync();
                backgroundWorker3.RunWorkerAsync();
            }
            else MessageBox.Show("Por favor confimar las filas de datos especificadas.");
        }

        private void startNamesCopy(NamesBlock names)
        {
            try
            {
                // Iteration through cells.
                for (int i = names.StartRow; i <= names.EndRow; i++)
                {
                    output.currentSheet.Cells[i, 6] = names.Section;
                    output.currentSheet.Cells[i, 7] = names.Group;
                    output.currentSheet.Cells[i, 8] = names.Category;
                    output.currentSheet.Cells[i, 9] = names.SubCategory;
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
                int maxRows = input.fullRange.Rows.Count;
                if (state == State.Testing) maxRows = 1000;
                int count = 0;


                int currentPosition = startingRows[0];
                int namesColumn = input.dataColumns[2];
                int quantityColumn = input.dataColumns[3];
                int sectionColumn = input.typeColumns[0];
                int groupColumn = input.typeColumns[0] + 1;
                int categoryColumn = input.typeColumns[0] + 2;
                int subCategoryColumn = input.typeColumns[0] + 3;
                string section = input.fullRange.Cells[startingRows[1], input.typeColumns[1]].Value2;
                string group = input.fullRange.Cells[startingRows[1] + 1, input.typeColumns[1] + 2].Value2;
                string category = input.fullRange.Cells[startingRows[1] + 2, input.typeColumns[1] + 4].Value2;
                string subCategory = input.fullRange.Cells[startingRows[1] + 3, input.typeColumns[1] + 6].Value2;


                // Iteration through cells.
                while (nullCount < 10 && workerThreadEnabled)
                {
                    if (input.fullRange.Cells[currentPosition, namesColumn] == null || input.fullRange.Cells[currentPosition, namesColumn].Value2 == null)
                    {
                        if (!dataCopied)
                        {
                            //Type copying
                            namesQ.Enqueue(new NamesBlock(startY, newSheetPositionY - 1, section, group, category, subCategory));

                            dataCopied = true;
                        }

                        int type = input.findUsedColumn(currentPosition, 1);
                        if (type == sectionColumn)
                        {
                            section = (input.fullRange.Cells[currentPosition, sectionColumn + 8]).Value2;
                        }
                        else if (type == groupColumn)
                        {
                            group = (input.fullRange.Cells[currentPosition, groupColumn + 9]).Value2;
                        }
                        else if (type == categoryColumn)
                        {
                            category = (input.fullRange.Cells[currentPosition, categoryColumn + 10]).Value2;
                        }
                        else if (type == subCategoryColumn)
                        {
                            subCategory = (input.fullRange.Cells[currentPosition, subCategoryColumn + 11]).Value2;
                        }
                        nullCount++;
                    }
                    else
                    {
                        if (dataCopied) startY = newSheetPositionY;
                        dataCopied = false;
                        nullCount = 0;

                        //Code copying.
                        output.currentSheet.Cells[newSheetPositionY, newSheetPositionX] = (input.fullRange.Cells[currentPosition, input.dataColumns[0]]).Value2;

                        //Name copying.
                        output.currentSheet.Cells[newSheetPositionY, newSheetPositionX + 1] = (input.fullRange.Cells[currentPosition, namesColumn]).Value2;

                        //Quantity copying.
                        output.currentSheet.Cells[newSheetPositionY, newSheetPositionX + 2] = (input.fullRange.Cells[currentPosition, quantityColumn]).Value2.ToString();

                        //Total copying.
                        output.currentSheet.Cells[newSheetPositionY, newSheetPositionX + 3] = (input.fullRange.Cells[currentPosition, quantityColumn + 2]).Value2.ToString();

                        //Type copying
                        //newSheet.Cells[newSheetPositionY, newSheetPositionX + 3] = section;
                        //newSheet.Cells[newSheetPositionY, newSheetPositionX + 4] = group;
                        //newSheet.Cells[newSheetPositionY, newSheetPositionX + 5] = category;
                        //newSheet.Cells[newSheetPositionY, newSheetPositionX + 6] = subCategory;

                        newSheetPositionY++;
                    }
                    currentPosition++;
                    count++;

                    //Test case
                    if (state == State.Testing)
                    {
                        backgroundWorker2.ReportProgress(100 * newSheetPositionY / maxRows);
                        if (newSheetPositionY >= 1000
                        && nullCount > 0)
                        {
                            workerThreadEnabled = false;
                        }
                    }

                    else if (count > 100)
                    {
                        backgroundWorker2.ReportProgress(100 * currentPosition / maxRows);
                        count = 0;
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
        /// Transfers a specified column on the Excel file into a new one, skipping null cells and filtering unwanted items.
        /// </summary>
        /// <param name="startingRow">Row to start parsing.</param>
        /// <param name="newSheetPositionX">Starting column cell in which the parsed data is copied.</param>
        /// <param name="newSheetPositionY">Starting row cell in which the parsed data is copied.</param>
        /// <param name="filterTerm">Filter term</param>
        private void startFiltering(int newSheetPositionX, int newSheetPositionY)
        {
            try
            {
                int nullCount = 0;
                int maxRows = input.fullRange.Rows.Count;
                if (state == State.Testing) maxRows = 1000;
                int count = 0;
                Dictionary<string, DataBlock> contents = new Dictionary<string, DataBlock>();


                int currentPosition = startingRows[0];
                int namesColumn = input.dataColumns[2];
                int quantityColumn = input.dataColumns[3];
                int sectionColumn = input.typeColumns[0];
                int groupColumn = input.typeColumns[0] + 1;
                int categoryColumn = input.typeColumns[0] + 2;
                int subCategoryColumn = input.typeColumns[0] + 3;
                string section = input.fullRange.Cells[startingRows[1], input.typeColumns[1]].Value2;
                string group = input.fullRange.Cells[startingRows[1] + 1, input.typeColumns[1] + 2].Value2;
                string category = input.fullRange.Cells[startingRows[1] + 2, input.typeColumns[1] + 4].Value2;
                string subCategory = input.fullRange.Cells[startingRows[1] + 3, input.typeColumns[1] + 6].Value2;


                // Iteration through cells.
                while (nullCount < 10)
                {
                    if (input.fullRange.Cells[currentPosition, namesColumn] == null || input.fullRange.Cells[currentPosition, namesColumn].Value2 == null)
                    {


                        int type = input.findUsedColumn(currentPosition, 1);
                        if (type == sectionColumn)
                        {
                            section = (input.fullRange.Cells[currentPosition, sectionColumn + 8]).Value2;
                        }
                        else if (type == groupColumn)
                        {
                            group = (input.fullRange.Cells[currentPosition, groupColumn + 9]).Value2;
                        }
                        else if (type == categoryColumn)
                        {
                            category = (input.fullRange.Cells[currentPosition, categoryColumn + 10]).Value2;
                        }
                        else if (type == subCategoryColumn)
                        {
                            subCategory = (input.fullRange.Cells[currentPosition, subCategoryColumn + 11]).Value2;
                        }
                        nullCount++;
                    }
                    else
                    {
                        nullCount = 0;
                        long code = Int64.Parse((input.fullRange.Cells[currentPosition, input.dataColumns[0]]).Value2);
                        string name = (input.fullRange.Cells[currentPosition, namesColumn]).Value2.ToUpper();

                        if (NamesList.Contains(name))
                        {
                            //Create DataBlock to store data.
                            DataBlock block;

                            if (!contents.TryGetValue(name, out block))
                            {
                                //Code copying.
                                output.currentSheet.Cells[newSheetPositionY, newSheetPositionX] = code.ToString();

                                //Name copying.

                                output.currentSheet.Cells[newSheetPositionY, newSheetPositionX + 1] = name;

                                //Quantity copying.
                                double quantity = (input.fullRange.Cells[currentPosition, quantityColumn]).Value2;
                                output.currentSheet.Cells[newSheetPositionY, newSheetPositionX + 2] = quantity.ToString();

                                //Total copying.
                                double total = (input.fullRange.Cells[currentPosition, quantityColumn + 2]).Value2;
                                output.currentSheet.Cells[newSheetPositionY, newSheetPositionX + 3] = total.ToString();

                                //Save in DataBlock

                                contents.Add(name, new DataBlock(name, code, quantity, total, newSheetPositionY));

                                //Unit Price copying.

                                /////////////////////////////////////////
                                //char unit = findUnitValue(name);
                                //switch (unit)
                                //{
                                //    case 'u':

                                //        output.currentSheet.Cells[newSheetPositionY, newSheetPositionX + 4] = "UNIDAD";
                                //        break;

                                //    case 'k':
                                //        output.currentSheet.Cells[newSheetPositionY, newSheetPositionX + 4] = "KG";
                                //        break;
                                //    case 'l':
                                //        output.currentSheet.Cells[newSheetPositionY, newSheetPositionX + 4] = "L";
                                //        break;
                                //    case 'n':
                                //        output.currentSheet.Cells[newSheetPositionY, newSheetPositionX + 4] = "n";
                                //        (output.currentSheet.Cells[newSheetPositionY, newSheetPositionX + 4] as Excel.Range).Interior.Color = Color.Orange;
                                //        break;
                                //    default:
                                //        (output.currentSheet.Cells[newSheetPositionY, newSheetPositionX + 4] as Excel.Range).Interior.Color = Color.Red;
                                //        break;
                                //}

                                output.currentSheet.Cells[newSheetPositionY, newSheetPositionX + 4] = getUnitaryPrice(name.ToUpper(), quantity, total);

                                /////////////////////////////////////////
                                //Type copying.
                                output.currentSheet.Cells[newSheetPositionY, newSheetPositionX + 5] = section;
                                output.currentSheet.Cells[newSheetPositionY, newSheetPositionX + 6] = group;
                                output.currentSheet.Cells[newSheetPositionY, newSheetPositionX + 7] = category;
                                output.currentSheet.Cells[newSheetPositionY, newSheetPositionX + 8] = subCategory;

                                //Date copying

                                output.currentSheet.Cells[newSheetPositionY, newSheetPositionX + 10] = input.Month;
                                output.currentSheet.Cells[newSheetPositionY, newSheetPositionX + 11] = input.Year;

                                newSheetPositionY++;
                            }
                            else
                            {
                                ///Modify existing cells to add data of duplicate cells;

                                //Quantity copying and saving in temp block.
                                double quantity = (input.fullRange.Cells[currentPosition, quantityColumn]).Value2 + block.Quantity;
                                block.Quantity = quantity;
                                output.currentSheet.Cells[block.DataRow, newSheetPositionX + 2] = quantity.ToString();

                                //Total copying and saving in temp block.
                                double total = (input.fullRange.Cells[currentPosition, quantityColumn + 2]).Value2 + block.Total;
                                block.Total = total;
                                output.currentSheet.Cells[block.DataRow, newSheetPositionX + 3] = total.ToString();

                                //Replace old DataBlock with the new one (update values).
                                contents.Remove(name);
                                contents.Add(name, block);
                                duplicatesProcessed++;

                            }
                        }
                    }
                    currentPosition++;
                    count++;

                    //Test case
                    if (state == State.Testing)
                    {
                        backgroundWorker2.ReportProgress(100 * newSheetPositionY / maxRows);
                        if (newSheetPositionY >= 1000
                        && nullCount > 0)
                        {
                            workerThreadEnabled = false;
                        }
                    }

                    else if (count > 100)
                    {
                        backgroundWorker2.ReportProgress(100 * currentPosition / maxRows);
                        count = 0;
                    }
                }
                foreach (string name in contents.Keys)
                {
                    if (NamesList.Contains(name))
                    {
                        NamesList.Remove(name);
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

        private char findUnitValue(string name)
        {
            name.ToUpper();
            int specialCase = -1;

            //Pre-tokenizing searches
            if (name.Contains("KILO"))
            {
                return 'k';
            }
            else if (name.Contains("PAQ") || name.Contains("PACK"))
            {

                if (name.Contains("UNIDADES"))
                {
                    specialCase = 1; // "UNIDADES" case
                }
                else if (name.Contains("UNIDAD"))
                {
                    specialCase = 3; // "UNIDAD" case
                }
                else
                {

                }
            }

            //String tokenizing
            double value;
            string[] words = name.Split(new char[] { ' ', '.', 'X', ',' });
            foreach (string word in words)
            {
                if (word.Contains("KG"))
                {
                    //Get value
                    if (word.Equals("KG"))
                    {
                        return 'k';
                    }
                    else
                    {

                        string temp = word.Replace("KG", "");
                        if (Double.TryParse(temp, out value)) return 'k';

                    }
                }

                else if (specialCase == 1 && Double.TryParse(word, out value))
                {
                    return 'p';
                }

                else if (word.Contains("UND"))
                {
                    //Get rid of UND to get amount.
                    string temp = word.Replace("UND", "");
                    if (Double.TryParse(temp, out value)) return 'u';
                }

                else if (word.Contains("GR"))
                {
                    string temp = word.Replace("GR", "");
                    if (Double.TryParse(temp, out value))
                    {
                        return 'k';
                    }
                }

                else if (word.Contains("UNIDAD"))
                {
                    if (word.Equals("UNIDAD"))
                    {
                        return 'u';
                    }
                    else
                    {
                        string temp = word.Replace("UNIDAD", "");
                        if (Double.TryParse(temp, out value)) return 'u';
                    }
                }

                else if (word.Contains("UN"))
                {
                    string temp = word.Replace("UN", "");
                    if (Double.TryParse(temp, out value))
                    {
                        return 'u';
                    }
                }

                else if (word.Contains("ML"))
                {
                    string temp = word.Replace("ML", "");
                    if (Double.TryParse(temp, out value))
                    {
                        return 'l';
                    }
                }

                else if (word.Contains("CC"))
                {
                    string temp = word.Replace("CC", "");
                    if (Double.TryParse(temp, out value))
                    {
                        return 'l';
                    }
                }

                else if (word.Contains("L"))
                {
                    //Get value
                    if (word.Equals("L"))
                    {
                        return 'l';
                    }
                    else
                    {

                        string temp = word.Replace("L", "");
                        if (Double.TryParse(temp, out value)) return 'l';

                    }
                }

                else if (word.Contains("K"))
                {
                    //Get value
                    if (word.Equals("K"))
                    {
                        return 'l';
                    }
                    else
                    {

                        string temp = word.Replace("L", "");
                        if (Double.TryParse(temp, out value)) return 'l';

                    }
                }

            }   //end foreach loop
            return 'n';
        }


        /// <summary>
        /// Background worker 1 work method.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BackgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            if (state == State.LoadingFile)
            {
                input = new ExcelData(fileName);
            }
        }

        /// <summary>
        /// Background worker 2 work method.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BackgroundWorker2_DoWork(object sender, DoWorkEventArgs e)
        {
            //In case the normal search button is pressed.
            timer.Start();
            if (state == State.FullProcessing || state == State.Testing)
            {
                int newSheetPositionX = 1;
                int newSheetPositionY = 3;
                startFullTransfer(newSheetPositionX, newSheetPositionY);
            }

            //In case the filter search button is pressed.
            else if (state == State.FilterProcessing)
            {
                //Initial data setup.
                int newSheetPositionX = 1;
                int newSheetPositionY = 3;

                //Start data filtering.
                startFiltering(newSheetPositionX, newSheetPositionY);
            }
            workerThreadEnabled = false;
        }

        /// <summary>
        /// Background worker 3 work method.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BackgroundWorker3_DoWork(object sender, DoWorkEventArgs e)
        {
            NamesBlock names;

            try
            {
                while (workerThreadEnabled || namesQ.Count > 0)
                {
                    if (namesQ.TryDequeue(out names))
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
        /// Updates the progress text.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Worker2_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            MainTextBox.Text = "Progreso: %" + e.ProgressPercentage;
        }

        /// <summary>
        /// Work completed event handler
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BackgroundWorkers_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (state == State.LoadingFile)
            {
                IsFileLoaded = true;
                AppLoadingImage.Visible = false;
                StartButton.Enabled = true;
                TestButton.Enabled = true;
                RowConfirmButton.Enabled = true;
                FilterStartButton.Enabled = true;
                OpenFileButton.Enabled = true;
                TestButton.Enabled = true;
                state = State.Idle;
            }

            else
            {
                activeThreads--;
                if (activeThreads < 1)
                {
                    StartButton.Enabled = true;
                    FilterStartButton.Enabled = true;
                    TestButton.Enabled = true;
                    LoadingImage.Visible = false;
                    input.dataColumnsReady = false;
                    input.typeColumnsReady = false;
                    //Displays the total time it took to carry out the processing.
                    if (timer.IsRunning)
                    {
                        timer.Stop();
                        TimeSpan ts = timer.Elapsed;
                        string totalTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}", ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds / 10);
                        MainTextBox.Text = "Tiempo total: " + totalTime + "\r\nDuplicados en archivo: " + duplicatesProcessed + "\r\nDuplicados en base: " + inputDuplicates;
                    }
                    RowBox1.BackColor = Color.White;
                    RowBox2.BackColor = Color.White;
                    output.excelApp.Visible = true;
                    timer.Reset();
                    duplicatesProcessed = 0;
                    inputDuplicates = 0;

                    // debug
                    StringBuilder sb = new StringBuilder();
                    int count = 1;
                    foreach (string missingCode in NamesList)
                    {
                        sb.Append(count++);
                        sb.Append(") ");
                        sb.Append(missingCode);
                        sb.Append("\r\n");
                    }
                    ListBox.Text = sb.ToString();
                    //

                    state = State.Idle;
                }
            }
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
                input.getColumns(Int32.Parse(RowBox1.Text), Int32.Parse(RowBox2.Text));
                startingRows = new int[2];
                startingRows[0] = Int32.Parse(RowBox1.Text);
                startingRows[1] = Int32.Parse(RowBox2.Text);
                input.dataColumnsReady = true;
                RowBox1.BackColor = Color.LightGreen;
                input.typeColumnsReady = true;
                RowBox2.BackColor = Color.LightGreen;
                MainTextBox.Text = input.fullRange.Rows.Count.ToString();
                input.Month = MonthTextBox.Text;
                input.Year = YearTextBox.Text;
            }
            catch (Exception ex)
            {
                input.dataColumnsReady = false;
                input.typeColumnsReady = false;
                RowBox1.BackColor = Color.White;
                RowBox2.BackColor = Color.White;
                MessageBox.Show("Error confirmando filas.");
            }
        }

        /// <summary>
        /// Actions to be carried out when the program is closing.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (IsFileLoaded) Cleanup();
        }
    }
}
