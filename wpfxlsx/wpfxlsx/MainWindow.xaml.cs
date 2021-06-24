using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Globalization;
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

namespace wpfxlsx
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        static string pathway = "";

        private void clickbut_Click(object sender, RoutedEventArgs e)
        {
            if ((pathway == "") & (pathtofile.Text == ""))
            {
                MessageBox.Show("Файл не выбран");
            }
            if ((pathway == "") & pathtofile.Text != "")
            {
                pathway = pathtofile.Text;
            }

            object rOnly = false;
            object SaveChanges = false;
            object MissingObj = System.Reflection.Missing.Value;
            string[,] m = new string[200, 200];

            Excel.Application app = new Excel.Application();

            Excel.Workbooks workbooks = null;
            Excel.Workbook workbook = null;
            Excel.Sheets sheets = null;
            Excel.Range UsedRange;
            Excel.Range urRows;
            Excel.Range urColums;
            double Start_of_week_balance = 0;
            double add_sales_taxes = 0;
            double Less_till_shorts = 0;
            double add_till_overs = 0;
            double less_safe_shorts = 0;
            double add_safe_overs = 0;
            double less_total_lodgments = 0;
            double less_petty_cash = 0;
            double add_change_order_received = 0;
            double adjust_eow_errors = 0;
            double End_of_week = 0;
            double Real_total = 0;
            double AOM_Variance = 0;
            int k = 0;
            int k1 = 0;
            int k2 = 0;
            int k3 = 0;
            int k4 = 0;
            int k5 = 0;
            int k6 = 0;
            int k7 = 0;
            int k8 = 0;
            int k9 = 0;
            int k10 = 0;
            try
            {
                workbooks = app.Workbooks;
                workbook = workbooks.Open(pathway, MissingObj, rOnly, MissingObj, MissingObj,
                                            MissingObj, MissingObj, MissingObj, MissingObj, MissingObj,
                                            MissingObj, MissingObj, MissingObj, MissingObj, MissingObj);

                // Получение всех страниц докуента
                sheets = workbook.Sheets;

                foreach (Excel.Worksheet worksheet in sheets)
                {
                    // Получаем диапазон используемых на странице ячеек
                    UsedRange = worksheet.UsedRange;
                    // Получаем строки в используемом диапазоне
                    urRows = UsedRange.Rows;
                    // Получаем столбцы в используемом диапазоне
                    urColums = UsedRange.Columns;

                    // Количества строк и столбцов
                    int RowsCount = urRows.Count;
                    int ColumnsCount = urColums.Count;
                    //ебать я тупой, я пытался получить значение, которое еще не прочиталось))))))))))))))))))))))
                    for (int i = 1; i <= RowsCount; i++)
                    {
                        for (int j = 1; j <= ColumnsCount; j++)
                        {
                            Excel.Range CellRange = (Excel.Range)UsedRange.Cells[i, j];
                            // Получение текста ячейки
                            string CellText = (CellRange == null || CellRange.Value2 == null) ? null :
                                                (CellRange as Excel.Range).Value2.ToString();
                            m[i, j] = CellText;

                        }
                    }

                    CultureInfo ci = (CultureInfo)CultureInfo.CurrentCulture.Clone();
                    ci.NumberFormat.NumberDecimalSeparator = ".";
                    // вот тут уже идет поиск
                    for (int i = 1; i <= RowsCount; i++)
                    {
                        for (int j = 1; j <= ColumnsCount; j++)
                        {
                            if (m[i, j] != null)
                            {

                                if (m[i, j] == "Start of week balance:")
                                {

                                    //Console.WriteLine($"Start of week balance: {m[i, j + 8]}");
                                    Start_of_week_balance = Convert.ToDouble(m[i, j + 8]);


                                }
                                if ((m[i, j] == "Total") & (j == 2))
                                {
                                    // Console.WriteLine($"Total Sales: {m[i, j + 1]}");
                                    add_sales_taxes = Convert.ToDouble(m[i, j + 1]);
                                }

                                if ((m[i, j] == "Total") & (j == 7))
                                {
                                    //   Console.WriteLine($"Less till shorts: {m[i, j + 7]}");
                                    //  Console.WriteLine($"Add till overs: {m[i, j + 9]}");
                                    //  Console.WriteLine($"Adjust EOW error {m[i, j + 13]}");


                                    Less_till_shorts = Convert.ToDouble(m[i, j + 7]);
                                    add_till_overs = Convert.ToDouble(m[i, j + 9]);
                                    adjust_eow_errors = Convert.ToDouble(m[i, j + 13]);
                                }

                                if ((m[i, j] == "Total") & (j == 32) & (k == 0))
                                {
                                    //    Console.WriteLine($"Less safe shorts: {m[i, j + 1]}");
                                    less_safe_shorts = Convert.ToDouble(m[i, j + 1]);
                                    //    Console.WriteLine($"Add safe overs: {m[i, j + 2]}");
                                    add_safe_overs = Convert.ToDouble(m[i, j + 2]);
                                    k = k + 1;
                                }
                                else if ((m[i, j] == "Total") & (j == 32) & (k != 0))
                                {
                                    //   Console.WriteLine($"Change order: {m[i, j + 1]}");
                                    add_change_order_received = Convert.ToDouble(m[i, j + 1]);
                                    //   Console.WriteLine($"Petty cash: {m[i, j + 2]}");
                                    less_petty_cash = Convert.ToDouble(m[i, j + 2]);
                                    k = k + 1;
                                }

                                if ((m[i, j] == "Total") & (j == 4))
                                {
                                    //    Console.WriteLine($"Less total lodgements: {m[i,j+22]}");

                                    less_total_lodgments = Convert.ToDouble(m[i, j + 22]);
                                    if (less_total_lodgments >= 0)
                                    {
                                        less_total_lodgments = (-1) * less_total_lodgments;
                                    }
                                    else
                                    {
                                        less_total_lodgments = less_total_lodgments;
                                    }
                                }

                                if (m[i, j] == "End of week balance:")
                                {

                                    //  Console.WriteLine($"End of week balance: {m[i, j + 8]}");
                                    End_of_week = Convert.ToDouble(m[i, j + 8]);
                                }


                            }
                        }
                    }
                    Real_total = Start_of_week_balance + add_sales_taxes + Less_till_shorts + add_till_overs + less_safe_shorts + add_safe_overs + less_total_lodgments +
                less_petty_cash + add_change_order_received + adjust_eow_errors;
                    Console.WriteLine($"Real Total: {Math.Round(Real_total, 2)}");
                    AOM_Variance = End_of_week - Real_total;



                    worksheet.get_Range("AO9", "AP21").Cells.Borders.Weight = Excel.XlBorderWeight.xlMedium;

                    Console.WriteLine($"AOM Variance: {Math.Round(AOM_Variance, 2)}");
                    worksheet.Cells[9, "AO"].Font.Size = 15;
                    worksheet.Cells[9, "AP"].Font.Size = 15;
                    worksheet.Cells[9, "AO"].Interior.Color = Excel.XlRgbColor.rgbCoral;
                    Console.WriteLine("тут зыписываю");
                    worksheet.Cells[9, "AO"] = "Start week balance";
                    worksheet.Cells[9, "AP"] = Start_of_week_balance;

                    worksheet.Cells[10, "AO"].Interior.Color = Excel.XlRgbColor.rgbCoral;
                    worksheet.Cells[10, "AO"].Font.Size = 15;
                    worksheet.Cells[10, "AP"].Font.Size = 15;
                    worksheet.Cells[10, "AO"] = "add sales including tax";
                    worksheet.Cells[10, "AP"] = add_sales_taxes;

                    worksheet.Cells[11, "AO"].Interior.Color = Excel.XlRgbColor.rgbCoral;
                    worksheet.Cells[11, "AO"].Font.Size = 15;
                    worksheet.Cells[11, "AP"].Font.Size = 15;
                    worksheet.Cells[11, "AO"] = "Less till shorts";
                    worksheet.Cells[11, "AP"] = Less_till_shorts;


                    worksheet.Cells[12, "AO"].Interior.Color = Excel.XlRgbColor.rgbCoral;
                    worksheet.Cells[12, "AO"].Font.Size = 15;
                    worksheet.Cells[12, "AP"].Font.Size = 15;
                    worksheet.Cells[12, "AO"] = "add till overs";
                    worksheet.Cells[12, "AP"] = add_till_overs;

                    worksheet.Cells[13, "AO"].Interior.Color = Excel.XlRgbColor.rgbCoral;
                    worksheet.Cells[13, "AO"].Font.Size = 15;
                    worksheet.Cells[13, "AP"].Font.Size = 15;
                    worksheet.Cells[13, "AO"] = "Less safe shorts";
                    worksheet.Cells[13, "AP"] = less_safe_shorts;

                    worksheet.Cells[14, "AO"].Interior.Color = Excel.XlRgbColor.rgbCoral;
                    worksheet.Cells[14, "AO"].Font.Size = 15;
                    worksheet.Cells[14, "AP"].Font.Size = 15;
                    worksheet.Cells[14, "AO"] = "add safe overs";
                    worksheet.Cells[14, "AP"] = add_safe_overs;

                    worksheet.Cells[15, "AO"].Interior.Color = Excel.XlRgbColor.rgbCoral;
                    worksheet.Cells[15, "AO"].Font.Size = 15;
                    worksheet.Cells[15, "AP"].Font.Size = 15;
                    worksheet.Cells[15, "AO"] = "Less total lodgments";
                    worksheet.Cells[15, "AP"] = less_total_lodgments;


                    worksheet.Cells[16, "AO"].Interior.Color = Excel.XlRgbColor.rgbCoral;
                    worksheet.Cells[16, "AO"].Font.Size = 15;
                    worksheet.Cells[16, "AP"].Font.Size = 15;
                    worksheet.Cells[16, "AO"] = "Less Petty cash (less if minus figure, add if positive figure)";
                    worksheet.Cells[16, "AP"] = less_petty_cash;

                    worksheet.Cells[17, "AO"].Interior.Color = Excel.XlRgbColor.rgbCoral;
                    worksheet.Cells[17, "AO"].Font.Size = 15;
                    worksheet.Cells[17, "AP"].Font.Size = 15;
                    worksheet.Cells[17, "AO"] = "add change order received";
                    worksheet.Cells[17, "AP"] = add_change_order_received;

                    worksheet.Cells[18, "AO"].Interior.Color = Excel.XlRgbColor.rgbCoral;
                    worksheet.Cells[18, "AO"].Font.Size = 15;
                    worksheet.Cells[18, "AP"].Font.Size = 15;
                    worksheet.Cells[18, "AO"] = "adjust for EOW error (if negative in WSSR then substract here, if positive then add)";
                    worksheet.Cells[18, "AP"] = adjust_eow_errors;


                    worksheet.Cells[19, "AO"].Font.Size = 15;
                    worksheet.Cells[19, "AP"].Font.Size = 15;
                    worksheet.Cells[19, "AO"] = "End Of Week balance(WSSR)";
                    worksheet.Cells[19, "AP"] = End_of_week;


                    worksheet.Cells[20, "AO"].Interior.Color = Excel.XlRgbColor.rgbCoral;
                    worksheet.Cells[20, "AO"].Font.Size = 15;
                    worksheet.Cells[20, "AP"].Font.Size = 15;
                    worksheet.Cells[20, "AO"] = "Real Total (результат вычислений по формуле)";
                    worksheet.Cells[20, "AP"] = Math.Round(Real_total, 2);

                    worksheet.Cells[21, "AO"].Font.Size = 15;
                    worksheet.Cells[21, "AP"].Font.Size = 15;
                    worksheet.Cells[21, "AO"] = "AOM Variance (End Of Week balance - Real Total)";
                    worksheet.Cells[21, "AP"] = Math.Round(AOM_Variance, 2);

                    workbook.Save();
                    // worksheet.Cells[1, "B"] = "Site";
                    //  worksheet.Cells[1, "C"] = "Cost";
                    // Очистка неуправляемых ресурсов на каждой итерации
                    if (urRows != null) Marshal.ReleaseComObject(urRows);
                    if (urColums != null) Marshal.ReleaseComObject(urColums);
                    if (UsedRange != null) Marshal.ReleaseComObject(UsedRange);
                    if (worksheet != null) Marshal.ReleaseComObject(worksheet);
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
            sowb.Content = Start_of_week_balance;
            asit.Content = add_sales_taxes;
            lts.Content = Less_till_shorts;
            ato.Content = add_till_overs;
            lss.Content = less_safe_shorts;
            aso.Content = add_safe_overs;
            ltl.Content = less_total_lodgments;
            pc.Content = less_petty_cash;
            eowb.Content = End_of_week;
            rt.Content = Math.Round(Real_total, 2);
            aomv.Content = Math.Round(AOM_Variance, 2);
            finalResult.Text = $"Start of week balance:\t {Start_of_week_balance}\nadd sales w/ taxes:\t {add_sales_taxes}\nLess till shorts:\t {Less_till_shorts}\nadd till overs:\t {add_till_overs}\nLess safe shorts:\t {less_safe_shorts}\n Add safe overs:\t {add_safe_overs}\nless total lodgements:\t {less_total_lodgments}\npetty cash:\t {less_petty_cash}\nadjust eow error:\t {adjust_eow_errors}\nend of week balance:\t {End_of_week}\nreal total:\t {Math.Round(Real_total, 2)}\nAOM variance:\t{Math.Round(AOM_Variance, 2)}";

            // pathway = pathtofile.Text;

        }

        private void choosefile_Click(object sender, RoutedEventArgs e)
        {

            OpenFileDialog ofd = new OpenFileDialog(); // создаём процесс  
            ofd.ShowDialog(); // открываем проводник    

            if (ofd.FileName != "") // проверка на выбор файла  
            {
                pathway = ofd.FileName;
            }
            else MessageBox.Show("Файл не выбран");



        }

    }
}
