using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.IO;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;

namespace SAWDocConverter
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void btn_LoadDoc_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                Parse780(openFileDialog1.FileName.ToString());
            }
        }

        private void Parse780(string fn)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;
            //Excel.Range range;

            xlApp = new Excel.Application();
            //xlWorkBook = xlApp.Workbooks.Open(@"C:\Users\zb024007\Documents\780.xlsx"); <-- alternate method
            xlWorkBook = xlApp.Workbooks.Open(fn);
            //xlWorkBook = xlApp.Workbooks.Open(@"d:\csharp-Excel.xls", 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0); <-- alternate method
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(4); //SAWSORT tab in 780

            int maxRows = 250; //max number of parameters
            int headerOffset = 4; //rows allocated for header
            char activeCol = 'A';

            List<string> numList = ExcelColtoList(xlWorkSheet, activeCol, maxRows, headerOffset, true);

            maxRows = numList.Count() + headerOffset; //update max rows




            //Test Code
            ExcelShowCellContents(xlWorkSheet, "E3"); //1st test type column is E3, 2nd column is H3...


            //Cleanup
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);

        }

        private static void ExcelShowCellContents(Excel.Worksheet xlWorkSheet, string cell)
        {
            System.Windows.MessageBox.Show(xlWorkSheet.get_Range(cell).Value2.ToString());
        }

        private static List<string> ExcelColtoList(Excel.Worksheet xlWorkSheet, char activeCol, int maxRows, int headerOffset, bool trimNulls)
        {
            int activeRow = headerOffset + 1;
            
            string activeCell = activeCol + activeRow.ToString();
            string cellContents = xlWorkSheet.get_Range(activeCell).Value2.ToString();
            List<string> numList = new List<string>();
            List<List<string>> fullList = new List<List<string>>();

            while (activeRow < maxRows)
            {
                activeCell = activeCol + activeRow.ToString();

                try
                {
                    //cellContents = xlWorkSheet.get_Range(activeCell).Value2.ToString();  <-- Does not work on null values
                    cellContents = Convert.ToString(xlWorkSheet.get_Range(activeCell).Value2);
                    numList.Add(cellContents);
                    activeRow++;
                }
                catch (Exception e1)
                {
                    System.Windows.MessageBox.Show(e1.ToString());
                }

            }

            if (trimNulls) { numList.RemoveAll(item => item == null); }

            System.Windows.MessageBox.Show(numList.Last());
            return numList;
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                System.Windows.MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

    }
}
