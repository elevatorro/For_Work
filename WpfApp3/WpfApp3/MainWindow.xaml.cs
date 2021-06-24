using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
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

namespace WpfApp3
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
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

            // pathway = pathtofile.Text;
            string[,] mass = new string[200, 200];
            int j = 0;
            string[,] newmass = new string[200, 200];
            String line = String.Empty;
            System.IO.StreamReader file = new System.IO.StreamReader(pathway);
            while ((line = file.ReadLine()) != null)
            {
                String[] parts_of_line = line.Split(',');
                for (int i = 0; i < parts_of_line.Length; i++)
                {
                    parts_of_line[i] = parts_of_line[i].Trim();
                    mass[j, i] = parts_of_line[i].Trim();
                }
                j++;
            }
            //  Console.WriteLine(mass.GetLength(0));

            //  Console.WriteLine(mass.GetLength(1));
            CultureInfo ci = (CultureInfo)CultureInfo.CurrentCulture.Clone();
            ci.NumberFormat.NumberDecimalSeparator = ".";
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


            for (int i = 0; i < mass.GetLength(0); i++)
            {
                for (j = 0; j < mass.GetLength(1); j++)
                {

                    //Console.Write(mass[i, j] + " ");
                    if (mass[i, j] == "Start of week balance:")
                    {
                        //Console.WriteLine(Convert.ToDouble(mass[i,j+1], ci)+2);
                        Start_of_week_balance = Convert.ToDouble(mass[i, j + 1], ci);
                        //  Console.WriteLine($"Start of week balance = {Start_of_week_balance}");
                        sowb.Content = Start_of_week_balance;
                    }


                    if (k == 0)
                    {
                        if (mass[i, j] == "Register Summary")
                        {
                            for (int z = j; z < mass.GetLength(1); z++)
                            {
                                if (mass[i, z] == "Total")
                                {
                                    k = k + 1;
                                    add_sales_taxes = Convert.ToDouble(mass[i, z + 1], ci);
                                    // Console.WriteLine($"Total sales w/ taxes = {add_sales_taxes}");
                                    asit.Content = add_sales_taxes;
                                }
                            }
                        }
                    }

                    if (k1 == 0)
                    {
                        if (mass[i, j] == "Register Summary")
                        {
                            for (int z = j; z < mass.GetLength(1); z++)
                            {
                                if (mass[i, z] == "Total")
                                {
                                    k1 = k1 + 1;
                                    Less_till_shorts = Convert.ToDouble(mass[i, z + 2], ci);
                                    // Console.WriteLine($"Less Till Shorts = {Less_till_shorts}");
                                    lts.Content = Less_till_shorts;
                                }
                            }
                        }
                    }

                    if (k2 == 0)
                    {
                        if (mass[i, j] == "Register Summary")
                        {
                            for (int z = j; z < mass.GetLength(1); z++)
                            {
                                if (mass[i, z] == "Total")
                                {
                                    k2 = k2 + 1;
                                    add_till_overs = Convert.ToDouble(mass[i, z + 3], ci);
                                    // Console.WriteLine($"Add Till Overs = {add_till_overs}");
                                    ato.Content = add_till_overs;
                                }
                            }
                        }
                    }

                    if (k3 == 0)
                    {
                        if (mass[i, j] == "Register Summary")
                        {
                            for (int z = j; z < mass.GetLength(1); z++)
                            {
                                if (mass[i, z] == "Total")
                                {
                                    k3 = k3 + 1;
                                    adjust_eow_errors = Convert.ToDouble(mass[i, z + 4], ci);
                                    Console.WriteLine($"Adjust EOW Errors = {adjust_eow_errors}");

                                }
                            }
                        }
                    }



                    if (k5 == 0)
                    {
                        if (mass[i, j] == "Safe Summary")
                        {
                            for (int z = j; z < mass.GetLength(1); z++)
                            {
                                if (mass[i, z] == "Total")
                                {
                                    k5 = k5 + 1;
                                    less_safe_shorts = Convert.ToDouble(mass[i, z + 1], ci);
                                    // Console.WriteLine($"Less Safe Shorts = {less_safe_shorts}");
                                    lss.Content = less_safe_shorts;
                                }
                            }
                        }
                    }

                    if (k4 == 0)
                    {
                        if (mass[i, j] == "Safe Summary")
                        {
                            for (int z = j; z < mass.GetLength(1); z++)
                            {
                                if (mass[i, z] == "Total")
                                {
                                    k4 = k4 + 1;
                                    add_safe_overs = Convert.ToDouble(mass[i, z + 2], ci);
                                    //  Console.WriteLine($"Add Safe Overs = {add_safe_overs}");
                                    aso.Content = add_safe_overs;
                                }
                            }
                        }
                    }

                    if (k6 == 0)
                    {
                        if (mass[i, j] == "Lodgements")
                        {
                            for (int z = j; z < mass.GetLength(1); z++)
                            {
                                if (k7 == 0)
                                {
                                    if ((mass[i, z] == "Gift Cards"))
                                    {
                                        k6 = k6 + 1;
                                        k7 = k7 + 1;
                                        z = z + 2;
                                        // less_total_lodgments = Convert.ToDouble(mass[i, z + 5], ci);
                                        // Console.WriteLine($"Lodgements = {less_total_lodgments}");
                                    }
                                }
                                if ((mass[i, z] == "Total") & (k7 == 1))
                                {
                                    k6 = k6 + 1;
                                    k7 = k7 + 1;
                                    less_total_lodgments = Convert.ToDouble(mass[i, z + 5], ci);
                                    if (less_total_lodgments <= 0)
                                    {
                                        Console.WriteLine($"Lodgements = {less_total_lodgments}");
                                        ltl.Content = less_total_lodgments;
                                    }

                                    else
                                    {
                                        less_total_lodgments = less_total_lodgments * (-1);
                                        Console.WriteLine($"Lodgements = {less_total_lodgments}");
                                        ltl.Content = less_total_lodgments;
                                    }
                                }
                            }
                        }
                    }

                    if (k8 == 0)
                    {
                        if (mass[i, j] == "Others")
                        {
                            for (int z = j; z < mass.GetLength(1); z++)
                            {
                                if (mass[i, z] == "Total")
                                {
                                    k8 = k8 + 1;
                                    less_petty_cash = Convert.ToDouble(mass[i, z + 2], ci);
                                    Console.WriteLine($"Less Safe Shorts = {less_petty_cash}");
                                    pc.Content = less_petty_cash;
                                }
                            }
                        }
                    }

                    if (k9 == 0)
                    {
                        if (mass[i, j] == "Others")
                        {
                            for (int z = j; z < mass.GetLength(1); z++)
                            {
                                if (mass[i, z] == "Total")
                                {
                                    k9 = k9 + 1;
                                    if (mass[i,z+1] == "")
                                    {
                                        add_change_order_received = 0;
                                        //add_change_order_received = Convert.ToDouble(mass[i, z + 1], ci);
                                    }
                                    else {
                                        add_change_order_received = Convert.ToDouble(mass[i, z + 1], ci);
                                    }
                                    Console.WriteLine($"Change order = {add_change_order_received}");

                                }
                            }
                        }
                    }

                    /* if (k10 == 0)
                     {
                         if (mass[i, j] == "End of week balance:")
                         {
                             for (int z = j; z < mass.GetLength(1); z++)
                             {
                                 if (mass[i, z] == "Total")
                                 {
                                     k10 = k10 + 1;
                                     End_of_week = Convert.ToDouble(mass[i, z + 1], ci);
                                     Console.WriteLine($"End Of Week Balance = {End_of_week}");
                                 }
                             }
                         }
                     }*/
                    if (mass[i, j] == "End of week balance:")
                    {
                        //Console.WriteLine(Convert.ToDouble(mass[i,j+1], ci)+2);
                        End_of_week = Convert.ToDouble(mass[i, j + 1], ci);
                        Console.WriteLine($"End Of Week Balance = {End_of_week}");
                        eowb.Content = End_of_week;
                    }



                }
                //  Console.Write('\n');


            }

            Real_total = Start_of_week_balance + add_sales_taxes + Less_till_shorts + add_till_overs + less_safe_shorts + add_safe_overs + less_total_lodgments +
                less_petty_cash + add_change_order_received + adjust_eow_errors;
            Real_total = Math.Round(Real_total,2);
            Console.WriteLine($"Real Total: {Math.Round(Real_total, 2)}");
            rt.Content = Math.Round(Real_total, 2);
            AOM_Variance = Math.Round(Real_total - End_of_week,2);
            Console.WriteLine($"AOM Variance: {Math.Round(AOM_Variance, 2)}");
            aomv.Content = Math.Round(AOM_Variance, 2);
            finalResult.Text = $"Start of week balance:\t {Start_of_week_balance}\nadd sales w/ taxes:\t {add_sales_taxes}\nLess till shorts:\t {Less_till_shorts}\nadd till overs:\t {add_till_overs}\nLess safe shorts:\t {less_safe_shorts}\n Add safe overs:\t {add_safe_overs}\nless total lodgements:\t {less_total_lodgments}\npetty cash:\t {less_petty_cash}\nadjust eow error:\t {adjust_eow_errors}\nend of week balance:\t {End_of_week}\nreal total:\t {Real_total}\nAOM variance:\t{AOM_Variance}";
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
