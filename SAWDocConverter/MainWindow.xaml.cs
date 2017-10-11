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
        public List<List<string>> activeSheet { get; private set; }
        public static Dictionary<Dictionary<string, int>, int> activeBookTabAndTestStepDict { get; private set; }
        public static string activeTestStep { get; private set; }
        public static string activeFilePath { get; private set; }
        public List<string> activeFileTabNames { get; private set; }
        public List<string> headerTemplate { get; private set; }

        public MainWindow()
        {
            InitializeComponent();
        }
        private void btn_CreateCSV_Click(object sender, RoutedEventArgs e)
        {
            string mask = string.Empty;

            foreach (string testTab in lbx_TabNames.SelectedItems)
            {
                activeTestStep = testTab.Replace("FREQ_", "");
                try
                {
                    mask = System.IO.Path.GetFileNameWithoutExtension(activeFilePath).Substring(0, 5); //assuming filename is same as saved in DocCenter
                }
                catch (System.ArgumentOutOfRangeException e1)
                {
                    System.Windows.MessageBox.Show(e1.ToString(), "Mask Number not contained in Filename"); //file too short
                    break;
                }

                updateHeaderTemplate(mask + "_" + activeTestStep);

                int index = 0;
                int offset = 0;

                foreach (var k in activeBookTabAndTestStepDict.Keys)
                {
                    if (k.ContainsKey(testTab))
                    {
                        index = activeBookTabAndTestStepDict[k];
                        k.TryGetValue(testTab, out offset);
                        //offset =
                        break;
                    }
                }

                activeSheet = Parse78x(activeFilePath, index, offset);

                StringBuilder sb = new StringBuilder();
                AddCSVJunk(activeSheet);
                FormatCSVtoFile(sb);
                string outputFilename = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                File.WriteAllText(outputFilename + mask + "_" + activeTestStep + " Limits.csv", sb.ToString());
            };
        }

        private void btn_LoadDoc_Click(object sender, RoutedEventArgs e)
        {
            using (OpenFileDialog openFileDialog1 = new OpenFileDialog
            {
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                Filter = "Excel Files (*.xlsx)|*.xlsx|All files (*.*)|*.*",
                FilterIndex = 1,
                RestoreDirectory = true
            }
            )
                if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    activeFilePath = openFileDialog1.FileName.ToString();
                    activeBookTabAndTestStepDict = Get78xTabsAndSteps(activeFilePath);


                    var temp = activeBookTabAndTestStepDict;

                    foreach (Dictionary<string, int> k in temp.Keys.ToList() )
                    {
                        if (k.Count == 0)
                        {
                            activeBookTabAndTestStepDict.Remove(k);
                        }
                    }

                    lbx_TabNames.Items.Clear();

                    foreach (Dictionary<string,int> dict1 in activeBookTabAndTestStepDict.Keys)
                    {
                        foreach(string testStep in dict1.Keys)
                        {
                            lbx_TabNames.Items.Add(testStep);
                        }
                    }

                }
                else
                {
                    System.Windows.MessageBox.Show("Failed to open File.");
                }
        }

        private Dictionary<Dictionary<string, int>, int> Get78xTabsAndSteps(string fn)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(fn);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            Dictionary<Dictionary<string, int>, int> tabsAndSteps = new Dictionary<Dictionary<string, int>, int>();

            /* Would like to be able to restrict number of tabs opened, but some documents in DocCenter contain the Instructions Tab
            int tabcount = 0;
            if (fn.Contains("780"))
            {
                tabcount = 6;
            }
            else if (fn.Contains("785"))
            {
                tabcount = 8;
            }
            */
            
            for (int i = 1; i < 8; i++) //max 8 tabs
            {
                try
                {
                    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(i);
                }
                catch (System.Runtime.InteropServices.COMException e)
                {
                    //System.Windows.MessageBox.Show(e.ToString()); --> End of File
                    break;
                }
                tabsAndSteps.Add(Get78xTestSteps(fn, i), i);
            };

            //tabNames.RemoveAll(x => x == null); --> Does not work for dictionaries, should be no null values if caught correctly above

            //Cleanup
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            ReleaseObject(xlWorkSheet);
            ReleaseObject(xlWorkBook);
            ReleaseObject(xlApp);

            return tabsAndSteps;
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
            activeFile.Insert(11, FillList(listSize, "0"));
            activeFile.Insert(12, FillList(listSize, ""));
            activeFile.Insert(13, FillList(listSize, ""));
            activeFile.Insert(14, FillList(listSize, "0"));
            activeFile.Insert(15, FillList(listSize, "0"));
        }

        public void updateHeaderTemplate(string step)
        {
            headerTemplate = new List<string>()
            {
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
                step + " LL",
                step + " UL"
            };

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
            string last = headerTemplate.Last();
            foreach (string s in headerTemplate)
            {
                if (s != last)
                {
                    sb.Append(s + ','); //end of line
                }
                else
                {
                    sb.Append(s);
                }
            }
            sb.AppendLine();

            for (int i = 0; i < activeSheet[0].Count; i++)
            {
                for (int j = 0; j < activeSheet.Count; j++)
                {
                    if (j == (activeSheet.Count - 1))
                    {
                        sb.Append(activeSheet[j][i]); //end of line
                    }
                    else
                    {
                        sb.Append(activeSheet[j][i] + ',');
                    }
                }

                sb.AppendLine();
            }
        }
        /// <summary>
        /// Obsolete, not being used
        /// </summary>
        /// <param name="fn"></param>
        /// <returns></returns>
        private Dictionary<string, int> Get78xTabNames(string fn)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(fn);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(  1);

            Dictionary<string, int> tabNames = new Dictionary<string, int>();

            //foreach (Excel.Worksheet worksheet in xlWorkBook) --> no def for GetEnumerator

            for (int i = 1; i < 8; i++)
            {
                try
                {
                    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(i);
                }
                catch (System.Runtime.InteropServices.COMException e)
                {
                    System.Windows.MessageBox.Show(e.ToString());
                    break;
                }
                tabNames.Add(Convert.ToString(xlWorkSheet.Name), i);
            };

            //tabNames.RemoveAll(x => x == null); --> Does not work for dictionaries, should be no null values if caught correctly above

            //Cleanup
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            ReleaseObject(xlWorkSheet);
            ReleaseObject(xlWorkBook);
            ReleaseObject(xlApp);

            return tabNames;
        }

        private Dictionary<string, int> Get78xTestSteps(string fn, int sheetIndex)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(fn);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(sheetIndex); // 4 = 1st tab (SAWSORT tab in 780 or IDT tab in 785)

            Dictionary<string, int> testStepNames = new Dictionary<string, int>();
            {
                if(Convert.ToString(xlWorkSheet.get_Range("E3").Value2) != null)
                {
                    testStepNames.Add(Convert.ToString(xlWorkSheet.get_Range("E3").Value2), 0);
                }
                if (Convert.ToString(xlWorkSheet.get_Range("H3").Value2) != null)
                {
                    testStepNames.Add(Convert.ToString(xlWorkSheet.get_Range("H3").Value2), 3);
                }
                if (Convert.ToString(xlWorkSheet.get_Range("K3").Value2) != null)
                {
                    testStepNames.Add(Convert.ToString(xlWorkSheet.get_Range("K3").Value2), 6);
                }
                if (Convert.ToString(xlWorkSheet.get_Range("N3").Value2) != null)
                {
                    testStepNames.Add(Convert.ToString(xlWorkSheet.get_Range("N3").Value2), 9);
                }
            }

            //RemoveNulls(testStepNames); --> no longer need?

            //Cleanup
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            ReleaseObject(xlWorkSheet);
            ReleaseObject(xlWorkBook);
            ReleaseObject(xlApp);

            return testStepNames;
        }

        private void RemoveNulls(Dictionary<string, int> dict)
        {
            throw new NotImplementedException();
            //dict = (from kv in dict
            //        where kv.Key != null
            //        select kv).ToDictionary(kv => kv.Key, kv => kv.Value);
        }

        private List<List<string>> Parse78x(string fn, int sheetIndex, int testStepIndexOffset)
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
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(sheetIndex); //1st tab is 4

            int maxRows = 250; //max number of parameters
            int headerOffset = 5; // 4 rows allocated for header + 1-based index for spreadsheet
            char activeCol = 'A';

            List<string> numList = ExcelColtoList(xlWorkSheet, activeCol, maxRows, headerOffset, true);
            List<List<string>> tableList = new List<List<string>>();
            tableList.Add(numList);

            maxRows = numList.Count() + headerOffset; //update max rows based on number of params found in first column, 1-based index

            for (int i = 1; i < 4; i++) //6 columns for param name and limits (4 for names and no limits)
            {
                //activeCol = (char)(activeCol + 1); //next column
                tableList.Add(ExcelColtoList(xlWorkSheet, ++activeCol, maxRows, headerOffset, false)); //increment column before using
            }

            tableList.Add(ExcelColtoList(xlWorkSheet, (char)(++activeCol + testStepIndexOffset), maxRows, headerOffset, false)); //1st Spec Col
            tableList.Add(ExcelColtoList(xlWorkSheet, (char)(activeCol + testStepIndexOffset + 1), maxRows, headerOffset, false)); //2nd Spec Col


            //Test Code
            //ExcelShowCellContents(xlWorkSheet, "E3"); //1st test type column is E3, 2nd column is H3...

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

        private void ListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            //need anything?
        }
    }
}
