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
        private readonly List<string> headerTemplate = new List<string> {
            "Enabled?",
            "Test Number",
            "Test Name",
            "Test Units",
            "Decimal Places",
            "Offset",
            "Failing Soft Bin",
            "Lower Guardband",
            "Upper Guardband",
            "BCS Enable",
            "Lower CU Tolerance",
            "Upper CU Tolerance",
            "Empty Socket LL",
            "Empty Socket UL",
            "PAT Lower Sigma Multiplier",
            "PAT Upper Sigma Multiplier",
            "85030_POSTPILLAR LL",
            "85030_POSTPILLAR UL" };

        public List<List<string>> activeFile { get; private set; }


        public MainWindow()
        {
            InitializeComponent();

        }
        private void btn_CreateCSV_Click(object sender, RoutedEventArgs e)
        {
            StringBuilder sb = new StringBuilder();

            AddCSVJunk(activeFile);
            FormatCSVtoFile(sb);

            File.WriteAllText("C:\\Users\\zb024007\\Documents\\SampleCSV.csv", sb.ToString());

        }

        private void btn_LoadDoc_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog
            {
                InitialDirectory = "C:\\Users\\zb024007\\Documents",
                Filter = "Excel Files (*.xlsx)|*.xlsx|All files (*.*)|*.*",
                FilterIndex = 1,
                RestoreDirectory = true
            };

            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                activeFile = Parse780(openFileDialog1.FileName.ToString());
            }
            else
            {
                System.Windows.MessageBox.Show("Failed to open File.");
            }
        }
        private void AddCSVJunk(List<List<string>> activeFile)
        {
            int listSize = activeFile[0].Count;

            activeFile.Insert(0, FillList(listSize, "YES"));
            activeFile.Insert(5, FillList(listSize, "0"));
            activeFile.Insert(6, FillList(listSize, "32"));
            activeFile.Insert(7, FillList(listSize, "0"));
            activeFile.Insert(8, FillList(listSize, "0"));
            activeFile.Insert(9, FillList(listSize, "0"));
            activeFile.Insert(10, FillList(listSize, "0"));
            activeFile.Insert(11, FillList(listSize, ""));
            activeFile.Insert(12, FillList(listSize, ""));
            activeFile.Insert(13, FillList(listSize, "0"));
            activeFile.Insert(14, FillList(listSize, "0"));
        }

        private static List<string> FillList(int size, string filler)
        {
            List<string> enabledList = new List<string>();
            for (int i = 0; i < size; i++)
            {
                enabledList.Add(filler);
            }

            return enabledList;
        }
        private void FormatCSVtoFile(StringBuilder sb)
        {
            foreach (string s in headerTemplate)
            {
                sb.Append(s + ',');
            }
            sb.AppendLine();

            for (int i = 0; i < activeFile[0].Count; i++)
            {
                for (int j = 0; j < activeFile.Count; j++)
                {
                    sb.Append(activeFile[j][i] + ',');
                }

                sb.AppendLine();
            }
        }
        private List<List<string>> Parse780(string fn)
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
            int headerOffset = 5; // 4 rows allocated for header + 1-based index for spreadsheet
            char activeCol = 'A';

            List<string> numList = ExcelColtoList(xlWorkSheet, activeCol, maxRows, headerOffset, true);
            List<List<string>> tableList = new List<List<string>>();
            tableList.Add(numList);

            maxRows = numList.Count() + headerOffset; //update max rows based on number of params found in first column, 1-based index

            for (int i = 1; i < 6; i++) //6 columns for param name and limits
            {
                activeCol = (char)(activeCol + 1); //next column
                tableList.Add(ExcelColtoList(xlWorkSheet, activeCol, maxRows, headerOffset, false));
            }

            //Test Code
            ExcelShowCellContents(xlWorkSheet, "E3"); //1st test type column is E3, 2nd column is H3...

            //Cleanup
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            ReleaseObject(xlWorkSheet);
            ReleaseObject(xlWorkBook);
            ReleaseObject(xlApp);

            return tableList;
        }

        private static void ExcelShowCellContents(Excel.Worksheet xlWorkSheet, string cell)
        {
            System.Windows.MessageBox.Show(xlWorkSheet.get_Range(cell).Value2.ToString());
        }

        private static List<string> ExcelColtoList(Excel.Worksheet xlWorkSheet, char activeCol, int maxRows, int headerOffset, bool trimNulls)
        {
            int activeRow = headerOffset; //1-based index

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

        private static void ReleaseObject(object obj)
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
