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
        public Main_Form()
        {
            InitializeComponent();
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
                excelApp.Visible = true;
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
            try
            {
                int i = Int32.Parse(cellBox1.Text);
                int j = Int32.Parse(cellBox2.Text);
                if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                {
                    DataTextBox.Text = (string)(xlRange.Cells[i, j] as Excel.Range).Value2; ;
                }
                else DataTextBox.Text = "Empty Cell";
            }

            catch (RuntimeBinderException theException)
            {
                int i = Int32.Parse(cellBox1.Text);
                int j = Int32.Parse(cellBox2.Text);
                DataTextBox.Text = ((double)(xlRange.Cells[i, j] as Excel.Range).Value2).ToString();
            }

            catch (Exception theException)
            {
                String errorMessage;
                errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, theException.Message);
                errorMessage = String.Concat(errorMessage, "\n Full String: ");
                errorMessage = String.Concat(errorMessage, theException.ToString());
                MessageBox.Show(errorMessage, "Error");
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
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
    }
}
