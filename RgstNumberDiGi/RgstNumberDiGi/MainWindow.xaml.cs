using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
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
using Excel = Microsoft.Office.Interop.Excel;

namespace RgstNumberDiGi
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        static string pathway = "";
        
        public MainWindow()
        {
            InitializeComponent();
        }



        private void choose_file_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog(); // создаём процесс  
            ofd.ShowDialog(); // открываем проводник    

            if (ofd.FileName != "") // проверка на выбор файла  
            {
                pathway = ofd.FileName;
            }
            else MessageBox.Show("Файл не выбран");
        }

        private void result1_Click(object sender, RoutedEventArgs e)
        {
            string finalresult = "";
            if (pathway == "")
            {
                MessageBox.Show("ti eblan?");
            }
            else {
                object rOnly = false;
                object SaveChanges = false;
                object MissingObj = System.Reflection.Missing.Value;
                string[] first = new string[10000];
                string[] second = new string[10000];
                string[] third = new string[10500];
                Excel.Application app = new Excel.Application();
                Excel.Workbooks workbooks = null;
                Excel.Workbook workbook = null;
                Excel.Sheets sheets = null;
                Excel.Range UsedRange;
                Excel.Range urRows;
                Excel.Range urColumns;
                try
                {
                    workbooks = app.Workbooks;
                    workbook = workbooks.Open(pathway, MissingObj, rOnly, MissingObj, MissingObj,
                                                MissingObj, MissingObj, MissingObj, MissingObj, MissingObj,
                                                MissingObj, MissingObj, MissingObj, MissingObj, MissingObj);

                    sheets = workbook.Sheets;

                    foreach (Excel.Worksheet worksheet in sheets)
                    {
                        UsedRange = worksheet.UsedRange;
                        urRows = UsedRange.Rows;
                        urColumns = UsedRange.Columns;

                        int RowsCount = urRows.Count;

                        int ColumnsCount = urColumns.Count;

                        for (int i = 1; i <= RowsCount; i++)
                        {
                            for (int j = 1; j <= ColumnsCount; j++)
                            {
                                Excel.Range CellRange = (Excel.Range)UsedRange.Cells[i, j];
                                // Получение текста ячейки
                                string CellText = (CellRange == null || CellRange.Value2 == null) ? null :
                                                    (CellRange as Excel.Range).Value2.ToString();
                                if (j == 1)
                                {
                                    first[i] = CellText;
                                    //  Console.Write($"f {first[i]} ");
                                }
                                if (j == 2)
                                {
                                    second[i] = CellText;
                                    //  Console.Write($"s {second[i]} ");
                                }
                            }
                            // Console.WriteLine();
                        }


                    }
                }

                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }

                finally
                {
                    /* Очистка оставшихся неуправляемых ресурсов */
                    if (sheets != null) Marshal.ReleaseComObject(sheets);
                    if (workbook != null)
                    {
                        workbook.Close(SaveChanges);
                        Marshal.ReleaseComObject(workbook);
                        workbook = null;
                    }

                    if (workbooks != null)
                    {
                        workbooks.Close();
                        Marshal.ReleaseComObject(workbooks);
                        workbooks = null;
                    }
                    if (app != null)
                    {
                        app.Quit();
                        Marshal.ReleaseComObject(app);
                        app = null;
                    }
                }



                third = second.Except(first).ToArray();
                for (int i = 1; i < third.Length; i++)
                {
                    if (i > 1)
                    {
                        finalresult = $"{finalresult}, {third[i]}";
                    }
                    else {
                        finalresult = $"{third[i]} ";
                    }
                    //Console.WriteLine(third[i]);
                }
                textbox1.Text = finalresult;
            }
        }
    }
}
