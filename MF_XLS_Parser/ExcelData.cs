using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Collections.Concurrent;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Xml;
using System.Windows.Forms;
using System.Reflection;

namespace MF_XLS_Parser
{
    public class ExcelData
    {

        /// <summary>
        /// Checks if data columns have been aquired.
        /// </summary>
        public bool dataColumnsReady = false;

        /// <summary>
        /// Checks if type columns have been aquired.
        /// </summary>
        public bool typeColumnsReady = false;

        /// <summary>
        /// Instance of the Excel application.
        /// </summary>
        public Excel.Application excelApp { get; set; }

        /// <summary>
        /// Workbook currently worked on.
        /// </summary>
        public Excel._Workbook currentWorkbook;

        /// <summary>
        /// Current sheet on the workbook.
        /// </summary>
        public Excel._Worksheet currentSheet;

        /// <summary>
        /// Full range of the worksheet.
        /// </summary>
        public Excel.Range fullRange;

        /// <summary>
        /// The data columns extracted from the excel file.
        /// </summary>
        public int[] dataColumns;

        /// <summary>
        /// The data type columns extracted from the excel file.
        /// </summary>
        public int[] typeColumns;

        public ExcelData(string fileName)
        {
            if(fileName is null)
            {
                excelApp = new Excel.Application();
                excelApp.Visible = true;
                currentWorkbook = (Excel._Workbook)(excelApp.Workbooks.Add(Missing.Value));
                currentSheet = (Excel._Worksheet)currentWorkbook.ActiveSheet;
            }
            else
            {
                excelApp = new Excel.Application();
                excelApp.Visible = false;
                currentWorkbook = excelApp.Workbooks.Open(fileName);
                currentSheet = (Excel.Worksheet)currentWorkbook.Worksheets.get_Item(1);
                fullRange = currentSheet.UsedRange;
            }
            dataColumnsReady = false;
            typeColumnsReady = false;
        }

        /// <summary>
        /// Analyzes the specified rows for the data and type columns.
        /// </summary>
        /// <param name="firstDataRow">The row where data begins.</param>
        /// <param name="firstTypeRow">The row where data types begin.</param>
        public void getColumns(int firstDataRow, int firstTypeRow)
        {
            int processedColumn = 0;
            int position = 1;
            dataColumns = new int[5];
            while (processedColumn < 5)
            {
                if ((fullRange.Cells[firstDataRow, position]).Value2 != null)
                {
                    if (processedColumn == 0 && !Double.TryParse((fullRange.Cells[firstDataRow, position].Value2.ToString()), out double result))
                    {
                        throw new Exception("Error encontrando primera columna de datos.");
                    }
                    dataColumns[processedColumn] = position;
                    processedColumn++;
                }
                position++;
                if (position >= 100) throw new Exception("No se encontraron columnas de datos.");
            }

            processedColumn = 0;
            position = 1;
            typeColumns = new int[4];
            while (processedColumn < 4)
            {
                if ((fullRange.Cells[firstTypeRow, position]).Value2 != null)
                {
                    typeColumns[processedColumn] = position;
                    processedColumn++;
                }
                position++;
                if (position >= 100) throw new Exception("No se encontraron columnas de datos.");
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
            while (fullRange.Cells[row, currentCol] == null || fullRange.Cells[row, currentCol].Value2 == null)
            {
                currentCol++;
                if (currentCol >= 70) return -1;
            }
            return currentCol;
        }
    }
}
