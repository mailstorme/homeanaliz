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
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using win = System.Windows;
using System.Diagnostics;
using Microsoft.Win32;
using System.IO;
using System.Globalization;
using System.Web;
using PdfSharp.Pdf.Printing;
using System.Diagnostics;


//using System.mscorlib;

namespace Подсчет_начислений
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        string[] R1C1 = new string[] { "0", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ", "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BK", "BL", "BM", "BN", "BO", "BP", "BQ", "BR", "BS", "BT", "BU", "BV", "BW", "BX", "BY", "BZ", "CA", "CB", "CC", "CD", "CE", "CF", "CG", "CH", "CI", "CJ", "CK", "CL", "CM", "CN", "CO", "CP", "CQ", "CR", "CS", "CT", "CU", "CV", "CW", "CX", "CY", "CZ", "DA", "DB", "DC", "DD", "DE", "DF", "DG", "DH", "DI", "DJ", "DK", "DL", "DM", "DN", "DO", "DP", "DQ", "DR", "DS", "DT", "DU", "DV", "DW", "DX", "DY", "DZ", "EA", "EB", "EC", "ED", "EE", "EF", "EG", "EH", "EI", "EJ", "EK", "EL", "EM", "EN", "EO", "EP", "EQ", "ER", "ES", "ET", "EU", "EV", "EW", "EX", "EY", "EZ" };
        string[] Date = new string[] {".04.2017", ".03.2017", ".02.2017", ".01.2017", ".12.2016", ".11.2016", ".10.2016", ".09.2016", ".08.2016", ".07.2016", ".06.2016", ".05.2016", ".04.2016", ".03.2016", ".02.2016", ".01.2016", };
        int esd;
        string Podkl;

        public MainWindow()
        {
            InitializeComponent();
            esd = 0;
        }


        public void CloseProcess(Process[] before)
        {
            Process[] List;
            List = Process.GetProcessesByName("EXCEL");
            foreach (Process proc in List)
            {
                if (!before.Contains(proc))
                    proc.Kill();
            }
        }


        private object[][] getarray(string path, int list,int[] columns) //возвращает массив указаных колонок 
        {
            #region Открытие Excel
            var ExcelApp = new Excel.Application();
            ExcelApp.Visible = false;
            Excel.Sheets excelsheets;
            Excel.Worksheet excelworksheet;
            //Excel.Workbooks workbooks;
            Excel.Workbook book;
            Excel.Range range = null;

            book = ExcelApp.Workbooks.Open(path, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            //book.ActiveSheet.get_Item(list);
            excelsheets = book.Worksheets;
            excelworksheet = (Excel.Worksheet)excelsheets.get_Item(list);

            #endregion
            Process[] List = Process.GetProcessesByName("EXCEL");

            int Rows = excelworksheet.UsedRange.Rows.Count;
            int Columns = excelworksheet.UsedRange.Columns.Count;
            object[][] arr = new object[columns.Length][];

            int icolumn = 0;
            foreach(int column in columns)
            {
                for (int i = 0; i < Columns + 1; i++)
                {
                    if (column == i)
                    {
                        object[,] massiv;
                        arr[icolumn] = new object[Rows];
                        range = excelworksheet.get_Range(R1C1[i] + "2:" + R1C1[i] + Rows.ToString());
                        massiv = (System.Object[,])range.get_Value(Type.Missing);
                        arr[icolumn] = massiv.Cast<object>().ToArray();
                        icolumn++;
                    }
                }
            }

            #region прошлый вариант взятия колонок
            /*
            int icount = 0;
            for (int i = 0; i < Columns + 1; i++)
                if (i == c1 || i == c2 || i == c3)
                {
                    object[,] massiv;
                    arr[icount] = new object[Rows - 1];
                    range = excelworksheet.get_Range(R1C1[i] + "2:" + R1C1[i] + Rows.ToString());
                    massiv = (System.Object[,])range.get_Value(Type.Missing);
                    arr[icount] = massiv.Cast<object>().ToArray();
                    icount++;
                }
            for (int i = 0; i < Columns + 1; i++)
                if (i == tarif)
                {
                    object[,] massiv;
                    arr[icount] = new object[Rows - 1];
                    range = excelworksheet.get_Range(R1C1[i] + "2:" + R1C1[i] + Rows.ToString());
                    massiv = (System.Object[,])range.get_Value(Type.Missing);
                    arr[icount] = massiv.Cast<object>().ToArray();
                    icount++;
                }
                */
#endregion

            #region Закрытие Excel

            book.Close(false,false,false);

            ExcelApp.Quit();

            
            ExcelApp = null;
            excelsheets = null;
            excelworksheet = null;
            book = null;
            range = null;
            #endregion
            CloseProcess(List);
            
            return arr;
        }


        private object[][] getbasearray(string path)
        {
            #region Открытие Excel
            var ExcelApp = new Excel.Application();
            ExcelApp.Visible = false;
            Excel.Sheets excelsheets;
            Excel.Worksheet excelworksheet;
            //Excel.Workbooks workbooks;
            Excel.Workbook book;
            Excel.Range range = null;

            book = ExcelApp.Workbooks.Open(path, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            //book.ActiveSheet.get_Item(1);
            excelsheets = book.Worksheets;
            excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);

            #endregion
            Process[] List = Process.GetProcessesByName("EXCEL");

            int Rows = excelworksheet.UsedRange.Rows.Count;
            int Columns = excelworksheet.UsedRange.Columns.Count;

            object[][] arr = new object[2][];

            int icount = 0;
            for (int i = 0; i < Columns + 1; i++)
                if (i == 1 || i == 2)
                {
                    object[,] massiv;
                    arr[icount] = new object[Rows - 1];
                    range = excelworksheet.get_Range(R1C1[i] + "2:" + R1C1[i] + Rows.ToString());
                    massiv = (System.Object[,])range.get_Value(Type.Missing);
                    arr[icount] = massiv.Cast<object>().ToArray();
                    icount++;
                }

            #region Закрытие Excel

            book.Close(false, false, false);

            ExcelApp.Quit();

            ExcelApp = null;
            excelsheets = null;
            excelworksheet = null;
            //workbooks = null;
            book = null;
            range = null;
            #endregion
            CloseProcess(List);

            return arr;
        }


        #region Выбор папки
        private string DirSelect()
        {
            FolderBrowserDialog DirDialog = new FolderBrowserDialog();
            DirDialog.Description = "Выбор директории";
            DirDialog.SelectedPath = @"C:\";
            DirDialog.ShowDialog();
            return DirDialog.SelectedPath;
        }
        #endregion

     


        

        public class diler
        {
            public object name;
            //public object kodTT;

            public int b;
            public int a;
            public int allincom;

            public int count1201;
            public int count1202;
            public int count1203;
            public int count12046;
            public int count120712;

            public int Tab;
            public int TabAll;
            public int Treg;
            public int TregAll;

            public int smart;
            public int yourCountry;
            public int allIn;
            public int warmWelc;
            public int smartAll;
            public int yourCountryAll;
            public int allInAll;
            public int warmWelcAll;


            public int secondmonth;

            public double sum;
            public double sumWithPred;


            public diler (object NAME, bool fir, bool sec, bool thi, bool from4, bool from7,object nachislenia,bool abonent,bool regula, bool abonentAll, bool regulaAll, bool SecondMonth, bool AllInBool, bool AllIn, bool SmartBool, bool Smart,bool YourCountryBool,bool YourCountry, bool WarmWelcBool, bool WarmWelc,double predel)
            {
                name = NAME;

                count1201 = (fir) ? 1 : 0;
                count1202 = (sec) ? 1 : 0;
                count1203 = (thi) ? 1 : 0;
                count12046 = (from4) ? 1 : 0;
                count120712 = (from7) ? 1 : 0;

                smart = (Smart) ? 1 : 0;
                yourCountry = (YourCountry) ? 1 : 0;
                smartAll = (SmartBool)? 1 : 0;
                yourCountryAll = (YourCountryBool) ? 1 : 0;
                allIn = (AllIn) ? 1 : 0;
                allInAll = (AllInBool) ? 1 : 0;
                warmWelc = (WarmWelc) ? 1 : 0;
                warmWelcAll = (WarmWelcBool) ? 1 : 0;

                secondmonth = (SecondMonth) ? 1 : 0;

                b = 0;
                a = 0;
                Tab = 0;
                TabAll = 0;
                Treg = 0;
                TregAll = 0;


                if (abonent)
                    Tab++;
                else if (regula)
                    Treg++;

                if (abonentAll)
                    TabAll++;
                else if (regulaAll)
                    TregAll++;

                allincom = 1;
                sum += Convert.ToDouble(nachislenia);
                sumWithPred += predel;

                
                
            }
        }



        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            win.MessageBox.Show(Combobox.Text);
            int period =  Convert.ToInt32(Period.Text);
            string[] DatePeriod = new string[period];
            for (int i = 0; i < period; i++)
            {
                DatePeriod[i] = Date[i];
            }

            string path1 = "";
            string path2 = "";
            if (Combobox.Text == "Megafon")
            {
                path1 = @"C:\Users\Andrey\Desktop\Мега абонентская.txt"; //мега
                path2 = @"C:\Users\Andrey\Desktop\Мега регулярная.txt";
            }
            if (Combobox.Text == "MTC")
            {
                path1 = @"C:\Users\Andrey\Desktop\МтсАбонент.txt";
                path2 = @"C:\Users\Andrey\Desktop\МтсРегуляр.txt";
            }
            string abonents = System.IO.File.ReadAllText(path1).Replace("\n", " ");
            string regular = System.IO.File.ReadAllText(path2).Replace("\n", " ");
            string NotFound = "";


            string ComisPath = new OpenExcelFile().Filenamereturn();
            if (ComisPath == "can not open file")
                return;

            int Rows;
            int Columns;

            #region Открытие Excel
            var ExcelApp = new Excel.Application();
            ExcelApp.Visible = false;
            Excel.Sheets excelsheets;
            Excel.Worksheet excelworksheet;
            //Excel.Workbooks workbooks;
            Excel.Workbook book;
            Excel.Range range;

            book = ExcelApp.Workbooks.Open(ComisPath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            //book.ActiveSheet.get_Item(1);
            excelsheets = book.Worksheets;
            excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);

            #endregion
            Process[] ExcelListBeforeStart = Process.GetProcessesByName("EXCEL");

            Rows = excelworksheet.UsedRange.Rows.Count;
            Columns = excelworksheet.UsedRange.Columns.Count;

            object[,] ComisAr = new object[Rows,Columns +1];

            range = excelworksheet.get_Range(R1C1[1] + "1:" + R1C1[Columns] + Rows.ToString());
            ComisAr = (System.Object[,])range.get_Value(Type.Missing);

            #region Закрытие Excel

            book.Close(false, false, false);

            ExcelApp.Quit();

            ExcelApp = null;
            excelsheets = null;
            excelworksheet = null;
            //workbooks = null;
            book = null;
            range = null;
            #endregion
            CloseProcess(ExcelListBeforeStart);


            string mas = "";
            for (int i = 1; i <= 10; i++)
            {
                for (int j = 1; j <= Columns; j++)
                {
                    mas += ComisAr[i, j] + " ";
                }
                mas += "\n";
            }


            List<diler> dilers = new List<diler>();
            
            for (int i = 2; i <= Rows; i++)
            {
                int PeriodNabludenia = Convert.ToInt32(ComisAr[i, 2]); 
                if (PeriodNabludenia > period)
                    continue;
                
                if (ComisAr[i, 1] == null || ComisAr[i, 1].ToString() == "" || ComisAr[i, 1].ToString() == " " || ComisAr[i, 1].ToString() == null)
                    continue;
                
                bool first = false;
                bool second = false;
                bool third = false;
                bool from4to6 = false;
                bool from7to12 = false;
                bool abonent = false;
                bool regula = false;

                bool abonentAll = false;
                bool regulaAll = false;

                bool SmartBool = false;
                bool YourCountryBool = false;
                bool WarmWelcBool = false;
                bool AllInBool = false;

                bool Smart = false;
                bool YourCountry = false;
                bool WarmWelc = false;
                bool AllIn = false;
                double predel = 0;


                if (abonents.Contains(ComisAr[i, 4].ToString()))
                {
                    abonentAll = true;

                    if (ComisAr[i, 4].ToString().Contains("Smart"))
                        SmartBool = true;

                    if (ComisAr[i, 4].ToString().Contains("Всё включено"))
                        AllInBool = true;
                }

                else if (regular.Contains(ComisAr[i, 4].ToString()))
                {
                    regulaAll = true;

                    if (ComisAr[i, 4].ToString().Contains("Тёплый приём"))
                        WarmWelcBool = true;

                    if (ComisAr[i, 4].ToString().Contains("Твоя страна"))
                        YourCountryBool = true;
                }
                else
                    if (!NotFound.Contains(ComisAr[i, 4].ToString()))
                    NotFound += ComisAr[i, 4].ToString() + " ;   ";


                double nach = Convert.ToDouble(ComisAr[i, 3]);

                if (nach >= 120)
                {
                    if (abonentAll)
                    {
                        abonent = true;

                        if (SmartBool)
                            Smart = true;

                        if (AllInBool)
                            AllIn = true;
                    }
                    else if (regulaAll)
                    {
                        regula = true;

                        if (WarmWelcBool)
                            WarmWelc = true;

                        if (YourCountryBool)
                            YourCountry = true;
                    }
                    else
                        if (!NotFound.Contains(ComisAr[i, 4].ToString()))
                            NotFound += ComisAr[i, 4].ToString() + " ;   ";



                    switch (PeriodNabludenia)
                    {
                        case 1:
                            first = true;
                            break;
                        case 2:
                            second = true;
                            break;
                        case 3:
                            third = true;
                            break;
                        case 4:
                            from4to6 = true;
                            break;
                        case 5:
                            from4to6 = true;
                            break;
                        case 6:
                            from4to6 = true;
                            break;
                        case 7: case 8: case 9: case 10: case 11: case 12:
                            from7to12 = true;
                            break;
                    }

                    if (Combobox.Text == "MTC" && !Smart)
                        predel = Nachisl(nach, first, second, third, from4to6, from7to12, 2);

                    if (Combobox.Text == "MTC" && Smart)
                        predel = Nachisl(nach, first, second, third, from4to6, from7to12, 1);

                    if (Combobox.Text == "Megafon")
                        predel = Nachisl(nach, first, second, third, from4to6, from7to12, 3);

                }

                bool SecondMonth = (PeriodNabludenia == 2) ? true : false;


                bool find = false;
                foreach (diler d in dilers)
                {
                    if (d.name.ToString() == ComisAr[i, 1].ToString())
                    {
                        d.sum += nach;
                        d.sumWithPred += predel;

                        if (abonent)
                            d.Tab++;
                        else if (regula)
                            d.Treg++;

                        if (abonentAll)
                            d.TabAll++;
                        else if (regulaAll)
                            d.TregAll++;

                        find = true;

                        d.allincom++;
                        if (first)
                            d.count1201++;
                        if (second)
                            d.count1202++;
                        if (third)
                            d.count1203++; 
                        if (from4to6)
                            d.count12046++;
                        if (from7to12)
                            d.count120712++;

                        if (SecondMonth) d.secondmonth++;
                        if (AllIn) d.allIn++;
                        if (AllInBool) d.allInAll++;
                        if (Smart) d.smart++;
                        if (YourCountry) d.yourCountry++;
                        if (SmartBool) d.smartAll++;
                        if (YourCountryBool) d.yourCountryAll++;
                        if (WarmWelc) d.warmWelc++;
                        if (WarmWelcBool) d.warmWelcAll++;

                        break;
                    }
                }

                if (!find)
                {
                    dilers.Add(new diler(ComisAr[i, 1],first,second,third,from4to6,from7to12,nach,abonent,regula,abonentAll,regulaAll,SecondMonth,AllInBool,AllIn,SmartBool,Smart,YourCountryBool,YourCountry,WarmWelcBool,WarmWelc,predel));
                }
            }

            win.MessageBox.Show("1-ый этап завершен. Укажите файл Базы");



            string BasePath = new OpenExcelFile().Filenamereturn();
            if (BasePath == "can not open file")
                return;

            basaseach(ref dilers, BasePath, DatePeriod,1,"b");

            win.MessageBox.Show("Конец 2.1-го этапа. Укажите файл Архивной базы");
                     

            BasePath = new OpenExcelFile().Filenamereturn();
            if (BasePath == "can not open file")
                return;

            basaseach(ref dilers, BasePath, DatePeriod, 1,"a");
            
            basaseach(ref dilers, BasePath, DatePeriod, 2, "a");
           
            basaseach(ref dilers, BasePath, DatePeriod, 3, "a");
            
            win.MessageBox.Show("Конец 2.4-го этапа. Укажите файл с точками для анализа");


            string toch = new OpenExcelFile().Filenamereturn();
            if (toch == "can not open file")
                return;
            object[][] tochki = getarray(toch,1, new int[] {3,1});

            int columnsinresult = 31;
            object[,] result = new object[dilers.Count, columnsinresult];

            int k = 0;
            int N = tochki[0].Length;

            for (int i = 0; i < N; i++)
            {
                string TT = "";
                if(tochki[1][i] != null)
                TT = (tochki[1][i].ToString() == null || tochki[1][i].ToString() == "" || tochki[1][i] == null) ? "" : tochki[1][i].ToString() + " - ";
                string DD = (tochki[0][i].ToString() == null) ? " " : tochki[0][i].ToString();
                string dilerName = TT + DD;

                foreach (diler d in dilers)
                    if (dilerName == d.name.ToString())
                    {
                        result[k, 0] = d.name;
                        result[k, 1] = d.b + d.a;
                        result[k, 2] = d.allincom;
                        result[k, 3] = d.sum;
                        result[k, 4] = d.count1201;
                        result[k, 5] = d.count1202;
                        result[k, 6] = d.count1203;
                        result[k, 7] = d.count12046;
                        result[k, 8] = d.count120712;
                        result[k, 9] = Math.Round((d.count1201 / Convert.ToDouble(d.allincom)),4);
                        result[k, 10] = Math.Round((d.count1202 / Convert.ToDouble(d.allincom)),4);
                        result[k, 11] = Math.Round((d.count1203 / Convert.ToDouble(d.allincom)),4);
                        result[k, 12] = Math.Round((d.count12046 / Convert.ToDouble(d.allincom)),4);
                        result[k, 13] = Math.Round((d.count120712 / Convert.ToDouble(d.allincom)),4);
                        result[k, 14] = Math.Round((d.sum / Convert.ToDouble(d.allincom)),0);
                        result[k, 15] = Math.Round((d.sum / Convert.ToDouble(d.a + d.b)),0);
                        result[k, 16] = Math.Round((Convert.ToDouble(d.count1201) / Convert.ToDouble(d.a + d.b)),4);
                        result[k, 17] = Math.Round(((d.count1201 + d.count1202 + d.count1203) / Convert.ToDouble(d.a + d.b)),4);
                        result[k, 18] = (((d.TabAll == 0) ? 0 : d.Tab / Convert.ToDouble(d.TabAll))).ToString("p") + "  (" + d.TabAll.ToString() + ")";
                        result[k, 19] = (((d.TregAll == 0) ? 0 : d.Treg / Convert.ToDouble(d.TregAll))).ToString("p") + "  (" + d.TregAll.ToString() + ")";


                        result[k, 20] = result[k, 18];
                        result[k, 21] = result[k, 19];
                        result[k, 22] = result[k, 10] = (Math.Round((d.count1202 / Convert.ToDouble(d.allincom)), 3)).ToString("P");
                        result[k, 23] = (Math.Round((d.count1202 / Convert.ToDouble(d.secondmonth)), 2)*100) + "% (" + d.secondmonth.ToString() + ")";
                        result[k, 24] = Math.Round((d.sum / Convert.ToDouble(d.allincom)), 0);

                        result[k, 25] = d.count1201.ToString() + ":" + d.count1202.ToString() + ":" + d.count1203.ToString() + ":" 
                            + (d.count12046 + d.count120712).ToString() + "  (" + d.allincom.ToString() + ")";
                        result[k, 26] = d.allincom.ToString() +" | "+ (d.b + d.a).ToString();
                        result[k, 27] = Math.Round(d.sum, 0).ToString() + " | " + Math.Round(d.sumWithPred,0).ToString();
                        result[k, 28] = Math.Round((d.sumWithPred / Convert.ToDouble(d.allincom)), 0);

                        result[k, 29] = 0;
                        if (Combobox.Text == "Megafon" && d.allInAll!=0)
                            result[k, 29] = Math.Round(d.allIn / Convert.ToDouble(d.allInAll), 2).ToString("P") + "  (" + d.allInAll.ToString() + ")";
                        if (Combobox.Text == "MTC" && d.smartAll !=0)
                            result[k, 29] = Math.Round(d.smart / Convert.ToDouble(d.smartAll), 2).ToString("P") + "  (" + d.smartAll.ToString() + ")";

                        result[k, 30] = 0;
                        if (Combobox.Text == "Megafon" && d.warmWelcAll != 0)
                            result[k, 30] = Math.Round(d.warmWelc / Convert.ToDouble(d.warmWelcAll), 2).ToString("P") + "  (" + d.warmWelcAll.ToString() + ")";
                        if (Combobox.Text == "MTC" && d.yourCountryAll != 0)
                            result[k, 30] = Math.Round(d.yourCountry / Convert.ToDouble(d.yourCountryAll), 2).ToString("P") + "  (" + d.yourCountryAll.ToString() + ")";
                        k++;
                        break;
                    }
            }


            string resPath = new OpenExcelFile().Filenamereturn();
            if (resPath == "can not open file")
                return;
            insert(resPath,result, dilers.Count,columnsinresult);

            win.MessageBox.Show(NotFound,"Конец программы");
        }


        public void insert(string path,object[,] arr,int rows, int col)
        {
            #region Открытие Excel
            var ExcelApp = new Excel.Application();
            ExcelApp.Visible = false;
            Excel.Sheets excelsheets;
            Excel.Worksheet excelworksheet;
            //Excel.Workbooks workbooks;
            Excel.Workbook book;
            Excel.Range range = null;

            book = ExcelApp.Workbooks.Open(path, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            //book.ActiveSheet.get_Item(1);
            excelsheets = book.Worksheets;
            excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);

            #endregion
            Process[] List = Process.GetProcessesByName("EXCEL");


            range = null;
            range = excelworksheet.get_Range(R1C1[1] + "1:" + R1C1[col] + "1");
            range.Value2 = new object[,] { {"Дилер Дистр", "Всего отгрузок за выбранный период" ,"Кол-во симок в комиссии", "Всего платежей" ,
                    "кол-во симок >120р в первом месяце" ,"кол-во симок >120р во втором месяце","кол-во симок >120р в третьем месяце",
             "кол-во симок >120р в 4-6 месяце"  ,"кол-во симок >120р в 7-12 месяце", "1) 1M" ,"2) 2M" ,"3) 3M" ,"4) 4-6M","5) 7-12M","6) платежи на комис" ,"7) платежи на отгрузки" ,
                    "8) хорошие (>120р) симки 1-го пер набл на кол-во отгрузок" ,"9) хорошие (>120р) симки 1,2,3 пер набл на кол-во отгрузок","C АП", "Без АП" } };

            range = null;
            range = excelworksheet.get_Range(R1C1[1] + "2:" + R1C1[col] + rows.ToString());
            range.Value2 = arr;


            #region Закрытие Excel

            book.Save();
            book.Close(false, false, false);

            ExcelApp.Quit();

            ExcelApp = null;
            excelsheets = null;
            excelworksheet = null;
            //workbooks = null;
            book = null;
            range = null;
            #endregion
            CloseProcess(List);
        }


        private int dateinper(string s)
        {
            s = s.Remove(0,3);
            int month = Convert.ToInt32(s.Substring(0,2));
            int per = (s.Contains("2016")) ? 5 + (12 - month) : 5 - month;
            return per;
        }


        // Создание Экселя Комиссии
        private void Button_Click_2(object sender, RoutedEventArgs e)
        {

            string ComisPath = new OpenExcelFile().Filenamereturn();
            if (ComisPath == "can not open file")
                return;


            if (Combobox.Text == "MTC")
            {
                object[][] ali = getarray(ComisPath,1,new int[] { 27, 95, 98, 19, 106 }); //MTC

                int N = ali[0].Length;
                object[,] ins = new object[N, 4];

                for (int i = 0; i < N; i++)
                {
                    string TT = (ali[4][i].ToString() == null || ali[4][i].ToString() == "" || ali[4][i] == null) ? "" : ali[4][i].ToString() + " - ";
                    string DD = (ali[1][i] == null) ? " " : ali[1][i].ToString();

                    ins[i, 0] = TT + DD;
                    ins[i, 2] = ali[0][i];
                    ins[i, 1] = ali[2][i];
                    ins[i, 3] = ali[3][i];
                }

                string resPath = new OpenExcelFile().Filenamereturn();
                if (resPath == "can not open file")
                    return;
                insert(resPath, ins, N, 4);
            }

            
            //МЕГА
            if (Combobox.Text == "Megafon")
            {
                object[][] ali = getarray(ComisPath,1, new int[] {10, 15, 48, 56, 66});

                int N = ali[0].Length;
                object[,] ins = new object[N, 4];

                for (int i = 0; i < N; i++)
                {
                    string TT = "";
                    if (ali[4][i] != null)
                        TT = ( ali[4][i].ToString() == "" ) ? "" : ali[4][i].ToString() + " - ";
                    string DD = (ali[3][i] == null) ? " " : ali[3][i].ToString();

                    ins[i, 0] = TT + DD;
                    ins[i, 1] = dateinper(ali[1][i].ToString());
                    ins[i, 2] = (ali[2][i] == null) ? 0 : ali[2][i];
                    ins[i, 3] = ali[0][i];
                }

                string resPath = new OpenExcelFile().Filenamereturn();
                if (resPath == "can not open file")
                    return;
                insert(resPath, ins, N, 4);
            }

            win.MessageBox.Show("Конец");
        }

        static private double Nachisl(double nach, bool m1, bool m2, bool m3, bool m4, bool m5,int k)
        {
            double Nach = 0;
            if (k == 1)
            {
                if (m1)
                    Nach = (nach > 4000) ? 4000 : nach;
                else if (m2)
                    Nach = (nach > 3000) ? 3000 : nach;
                else if (m3)
                    Nach = (nach > 2500) ? 2500 : nach;
                else if (m4)
                    Nach = (nach > 1500) ? 1500 : nach;
                else if (m5)
                    Nach = (nach > 1000) ? 1000 : nach;
            }

            if (k == 2)
            {
                if (m1)
                    Nach = (nach > 2000) ? 2000 : nach;
                else if (m2)
                    Nach = (nach > 1500) ? 1500 : nach;
                else if (m3)
                    Nach = (nach > 800) ? 800 : nach;
                else if (m4)
                    Nach = (nach > 600) ? 600 : nach;
                else if (m5)
                    Nach = (nach > 500) ? 500 : nach;
            }

            if (k == 3)
            {
                if (m1)
                    Nach = (nach > 6000) ? 6000 : nach;
                else if (m2)
                    Nach = (nach > 5500) ? 5500 : nach;
                else if (m3)
                    Nach = (nach > 4000) ? 4000 : nach;
                else if (m4)
                    Nach = (nach > 3000) ? 3000 : nach;
                else if (m5)
                    Nach = (nach > 3000) ? 3000 : nach;
            }
            return Nach;
        }

        public void basaseach(ref List<diler> dilers,string BasePath,string[] DatePeriod, int list,string a)
        {
            object[][] basearr = getarray(BasePath, list, new int[] { 2, 11, 10, 18, 15 });
            int Nbase = basearr[0].Length;

            //win.MessageBox.Show(basearr[4][137].ToString());

            for (int i = 0; i < Nbase; i++)
            {
                if (basearr[1][i] == null)
                    continue;
                if (Combobox.Text == "MTC")
                {
                    if ((basearr[4][i].ToString() != "МТС" && basearr[4][i].ToString() != "МТС"))
                        continue;
                }
                if (Combobox.Text == "Megafon")
                {
                    if (!basearr[4][i].ToString().Contains("Мфон Дилерский") && !basearr[4][i].ToString().Contains("Мфон Дил ЗФ"))
                        continue;
                }
                if (basearr[2][i] == null && basearr[3][i] == null)
                {
                    continue;
                }
                if (basearr[1][i] == null)
                    continue;

                foreach (string date in DatePeriod)
                {
                    if (basearr[1][i].ToString().Contains(date))
                    {
                        string TT = "";
                        if (basearr[3][i] != null)
                            TT = (basearr[3][i].ToString() == "") ? "" : basearr[3][i].ToString() + " - ";
                        string DD = (basearr[2][i] == null) ? " " : basearr[2][i].ToString();

                        foreach (diler d in dilers)
                        {
                            if (d.name.ToString().Contains(TT + DD))
                                if (a == "a")
                                    d.a++;
                                else if (a == "b")
                                    d.b++;
                        }
                        break;
                    }
                }
            }
        }


    }
}
