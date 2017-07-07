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
        string[] Date = new string[] {".05.2017", ".04.2017", ".03.2017", ".02.2017", ".01.2017", ".12.2016", ".11.2016", ".10.2016", ".09.2016", ".08.2016", ".07.2016", ".06.2016", ".05.2016", ".04.2016", ".03.2016", ".02.2016", ".01.2016", };
        int TariffsCheck;
        string Podkl;

        public MainWindow()
        {
            InitializeComponent();
        }


        public void CloseProcess(Process[] before) //закрытие массива процессов (для закрытия процессов EXCEL) 
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


        private object[,] getarray(string path, ref int Rows)
        {
            int Columns;

            #region Открытие Excel
            var ExcelApp = new Excel.Application();
            ExcelApp.Visible = false;
            Excel.Sheets excelsheets;
            Excel.Worksheet excelworksheet;
            //Excel.Workbooks workbooks;
            Excel.Workbook book;
            Excel.Range range;

            book = ExcelApp.Workbooks.Open(path, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            //book.ActiveSheet.get_Item(1);
            excelsheets = book.Worksheets;
            excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);

            #endregion
            Process[] ExcelListBeforeStart = Process.GetProcessesByName("EXCEL");

            Rows = excelworksheet.UsedRange.Rows.Count;
            Columns = excelworksheet.UsedRange.Columns.Count;

            object[,] ComisAr = new object[Rows, Columns + 1];

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


            return ComisAr;
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



        #region ПЕРЕПИСЫВАЮ ЦИКЛ ПРЕОБРАЗОВАНИЯ В СПИСОК

        public List<diler> ListCreate(object[,] ComisAr, int period, ref string NotFound, string[] tariffs, int between)
        {
            List<diler> dilers = new List<diler>();
            int N = ComisAr.GetLength(0);
            int K = tariffs.Length;

            for (int i = 2; i <= N; i++)
            {
                if (ComisAr[i, 1] == null || ComisAr[i, 1].ToString() == "" || ComisAr[i, 1].ToString() == " " || ComisAr[i, 1].ToString() == null)
                    continue;

                int PeriodNabludenia = Convert.ToInt32(ComisAr[i, 2]);
                if (PeriodNabludenia > period)
                    continue;



                double nach = Convert.ToDouble(ComisAr[i, 3]);

                bool findtarif = false;
                int tariffID = 0;
                for (int j = 0; j < K; j++)
                {
                    if (tariffs[j] == ComisAr[i, 4].ToString())
                    {
                        tariffID = j;
                        findtarif = true;
                        break;
                    }
                }
                if (!findtarif)
                {
                    TariffsCheck++;

                    if (!NotFound.Contains(ComisAr[i, 4].ToString()))
                        NotFound += ComisAr[i, 4].ToString() + " ;   ";
                }


                bool find = false;
                foreach (diler d in dilers)
                {
                    if (d.name.ToString() == ComisAr[i, 1].ToString())
                    {
                        string date = ComisAr[i,6].ToString();
                        d.AddCom(PeriodNabludenia, tariffID, nach, Combobox.Text,date);
                        find = true;
                        break;
                    }
                }
                if (!find)
                {
                    dilers.Add(new diler(ComisAr[i, 1],tariffs,between,ComisAr[i,5]));
                    int listInt = dilers.Count -1;
                    string date = ComisAr[i, 6].ToString();
                    dilers[listInt].AddCom(PeriodNabludenia, tariffID,nach, Combobox.Text,date);
                }
            }
            return dilers;
        }
        #endregion



        public class diler
        {
            public object name;

            public int inBase;
            public int inArhiv;
            public int allincom;

            public double sum;
            public double sumWithPred;

            public object priznakCom;

            // новые переменные
            public tarifInfo[] tariffs;
            int between;

            public UserCount users;
            

            public diler (object NAME, string[] Tariffs, int Between, object priznakcom)
            {
                name = NAME;

                priznakCom = priznakcom;
                between = Between;
                tariffs = new tarifInfo[Tariffs.Length];

                int i = 0;
                foreach (string t in Tariffs)
                {
                    tariffs[i] = new tarifInfo(t);
                    i++;
                }

                inBase = 0;
                inArhiv = 0;
                allincom = 0;
                sum = 0;
                sumWithPred = 0;

                users = new UserCount();
            }

            public void AddCom(int period, int tariffID, double nach, string combobox, string date)
            {
                allincom++;
                double predel;

                if (combobox == "Megafon" && period == 2)
                {
                    double d = 61 - Convert.ToDouble(date.Substring(0, 2));
                    double dt = d / 61;
                    nach = nach * dt;
                }


                if (nach >= 120)
                {
                    users.AddGoodUser(period);
                    tariffs[tariffID].goodCount++;
                    predel = Predel(nach,tariffs[tariffID].tarif, period);
                }
                else
                {
                    users.AddBadUser(period);
                    tariffs[tariffID].count++;
                    predel = nach;
                }

                sum += nach;
                sumWithPred += predel;
            }

            public int AllGoodAbTariffs()
            {
                int count = 0;
                for (int i = 0; i < between; i++)
                {
                    count += tariffs[i].goodCount;
                }
                return count;
            }

            public int AllAbTariffs()
            {
                int count = 0;
                for (int i = 0; i < between; i++)
                {
                    count += tariffs[i].count;
                }
                return count;
            }

            public int AllGoodRegTariffs()
            {
                int count = 0;
                for (int i = between; i < tariffs.Length; i++)
                {
                    count += tariffs[i].goodCount;
                }
                return count;
            }

            public int AllRegTariffs()
            {
                int count = 0;
                for (int i = between; i < tariffs.Length; i++)
                {
                    count += tariffs[i].count;
                }
                return count;
            }

            public string TariffStatistic(string[] SelectTariff)
            {
                int c = 0;
                int gc = 0;
                string str = "";
                foreach(tarifInfo t in tariffs)
                {
                    foreach(string s in SelectTariff)
                    if (t.tarif.Contains(s))
                    {
                        c += t.count;
                        gc += t.goodCount;
                    }
                }
                str = (c != 0) ? Math.Round(gc / Convert.ToDouble(c), 2).ToString("P") + " (" + c + ")" : (0).ToString("P") + " (" + c + ")";

                return str;
            }
        }



        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            TariffsCheck = 0;

            
            if (Combobox.SelectedIndex == -1)
            {
                win.MessageBox.Show("Укажите оператора", "Ошибка");
                return;
            }
            int period = (Period.Text != "") ? Convert.ToInt32(Period.Text) : 0 ;
            if (period == 0)
            {
                win.MessageBox.Show("Укажите период", "Ошибка");
                return;
            }
            string[] DatePeriod = new string[period];
            for (int i = 0; i < period; i++)
            {
                DatePeriod[i] = Date[i];
            }

            string path1 = "";
            string path2 = "";
            if (Combobox.Text == "Megafon")
            {
                path1 = @"C:\Users\Andrei\Desktop\Тарифы\Мега абонентская.txt"; // мега
                path2 = @"C:\Users\Andrei\Desktop\Тарифы\Мега регулярная.txt"; // мтс
            }
            if (Combobox.Text == "MTC")
            {
                path1 = @"C:\Users\Andrei\Desktop\Тарифы\МтсАбонент.txt";
                path2 = @"C:\Users\Andrei\Desktop\Тарифы\МтсРегуляр.txt";
            }

            win.MessageBox.Show("Укажите файл комиссии");

            string[] AbArr  = System.IO.File.ReadAllLines(path1);
            string[] RegArr = System.IO.File.ReadAllLines(path2);
            string[] AllTariffs = AbArr.Union(RegArr).ToArray();
            string NotFound = "";

            string ComisPath = new OpenExcelFile().Filenamereturn();
            if (ComisPath == "can not open file")
                return;
            int Rows = 0;
            object[,] ComisAr = getarray(ComisPath,ref Rows);


            List<diler> dilers = ListCreate(ComisAr,period, ref NotFound,AllTariffs,AbArr.Length);


            win.MessageBox.Show("1-ый этап завершен. Укажите файл Базы  " + dilers.Count.ToString());



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

            int columnsinresult = 26;  //////////////////////////////////////////////////////////////

            object[,] result = new object[dilers.Count, columnsinresult];

            int k = 0;
            int N = tochki[0].Length;

            int count = 0;
            for (int i = 0; i < N; i++)
            {
                string TT = "";
                //if(tochki[1][i] != null)
                //TT = (tochki[1][i].ToString() == null || tochki[1][i].ToString() == "" || tochki[1][i] == null) ? "" : tochki[1][i].ToString() + " - ";
                string DD = (tochki[0][i].ToString() == null) ? " " : tochki[0][i].ToString();
                string dilerName = DD;
                

                foreach (diler d in dilers)
                    if (dilerName == d.name.ToString())
                    {
                        int col = 0;
                        int Otgruz = d.inBase + d.inArhiv;
                        int GoodAbTariffs = d.AllGoodAbTariffs();
                        int AllAbTarifs = d.AllAbTariffs();
                        int GoodRegTariffs = d.AllGoodRegTariffs();
                        int AllRegTariffs = d.AllRegTariffs();

                        result[k, col++] = d.name;
                        result[k, col++] = d.sum;
                        result[k, col++] = d.users.gfirst + d.users.gsecond + d.users.gthird + d.users.gm46 + d.users.gm79;
                        result[k, col++] = d.allincom;
                        result[k, col++] = Math.Round((d.users.gfirst / Convert.ToDouble(d.allincom)),4);
                        result[k, col++] = Math.Round((d.users.gsecond / Convert.ToDouble(d.allincom)),4);
                        result[k, col++] = Math.Round((d.users.gthird / Convert.ToDouble(d.allincom)),4);
                        result[k, col++] = Math.Round((d.users.gm46 / Convert.ToDouble(d.allincom)),4);
                        result[k, col++] = Math.Round((d.users.gm79 / Convert.ToDouble(d.allincom)),4);
                        result[k, col++] = Math.Round((d.sum / Convert.ToDouble(d.allincom)),0);
                        result[k, col++] = (Otgruz != 0)?Math.Round((d.sum / Convert.ToDouble(Otgruz)),0):0;
                        result[k, col++] = (Otgruz != 0)?Math.Round((Convert.ToDouble(d.users.gfirst) / Convert.ToDouble(Otgruz)),4) : 0;
                        result[k, col++] = (Otgruz != 0)?Math.Round(((d.users.gfirst + d.users.gsecond + d.users.gthird) / Convert.ToDouble(Otgruz)),4): 0;

                        result[k, col++] = (d.users.first != 0) ? (Math.Round((d.users.gfirst / Convert.ToDouble(d.users.first)), 1) * 100) + "% (" + d.users.first.ToString() + ")" : 0.ToString("P") + " (0)";
                        result[k, col++] = (d.users.second != 0) ? (Math.Round((d.users.gsecond / Convert.ToDouble(d.users.second)), 1) * 100) + "% (" + d.users.second.ToString() + ")" : 0.ToString("P") + " (0)";
                        result[k, col++] = (d.users.third != 0) ? (Math.Round((d.users.gthird / Convert.ToDouble(d.users.third)), 1) * 100) + "% (" + d.users.third.ToString() + ")" : 0.ToString("P") + " (0)";

                        result[k, col++] = "";

                        result[k, col++] = (d.users.second!=0)?(Math.Round((d.users.gsecond / Convert.ToDouble(d.users.second)), 1) * 100) + "% (" + d.users.second.ToString() + ")": 0.ToString("P")+" (0)" ;

                        result[k, col++] = (((AllAbTarifs == 0) ? 0 : GoodAbTariffs / Convert.ToDouble(AllAbTarifs))).ToString("p") + "  (" + AllAbTarifs.ToString() + ")";
                        result[k, col++] = (((AllRegTariffs == 0) ? 0 : GoodRegTariffs / Convert.ToDouble(AllRegTariffs))).ToString("p") + "  (" + AllRegTariffs.ToString() + ")";

                        result[k, col++] = Math.Round((d.sumWithPred / Convert.ToDouble(d.allincom)), 0);

                        result[k, col++] = (Combobox.Text != "Megafon") ? d.TariffStatistic(new string[] { "Smart", "smart" }) : d.TariffStatistic(new string[] { "Всё включено", "всё включено" });

                        result[k, col++] = (Combobox.Text != "Megafon") ? d.TariffStatistic(new string[] { "Твоя страна" }) : d.TariffStatistic(new string[] { "Тёплый приём", "тёплый приём" });

                        result[k, col++] = d.users.gfirst.ToString() + ":" + d.users.gsecond.ToString() + ":" + d.users.gthird.ToString() + ":" 
                            + (d.users.gm46 + d.users.gm79).ToString() + "  (" + d.allincom.ToString() + ")";

                        result[k, col++] = d.allincom.ToString() +" | "+ (Otgruz).ToString();

                        result[k, col++] = d.priznakCom;

                        count = col;
                        k++;
                        break;
                    }
            }
            

            string resPath = new OpenExcelFile().Filenamereturn();
            if (resPath == "can not open file")
                return;
            insert(resPath,result, dilers.Count,count);

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
            range.Value2 = new object[,] { {"Дилер Дистр", "Всего платежей" ,"Хорошие симки за период", "Всего симок в комиссии","кол-во симок >120р в первом месяце" ,"кол-во симок >120р во втором месяце","кол-во симок >120р в третьем месяце",
             "кол-во симок >120р в 4-6 месяце"  ,"кол-во симок >120р в 7-12 месяце", "6) платежи на комис" ,"7) платежи на отгрузки" ,
                    "8) хорошие (>120р) симки 1-го пер набл на кол-во отгрузок" ,"9) хорошие (>120р) симки 1,2,3 пер набл на кол-во отгрузок","1м активность","2м активность","3м активность","Подобие","2м активность","тарифы с АП",
                    "тарифы без АП","Среднее пополнение","Тариф 1","Тариф 2","1м:2м:3м:4+м (ком)","в комиссии | отгрузки за период" } };

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
            int per = (s.Contains("2016")) ? 6 + (12 - month) : 6 - month;
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
                object[][] ali = getarray(ComisPath,1,new int[] { 27, 95, 98, 19, 106, 96 }); //MTC

                int N = ali[0].Length;
                object[,] ins = new object[N, 6];

                for (int i = 0; i < N; i++)
                {
                    string TT = (ali[4][i].ToString() == null || ali[4][i].ToString() == "" || ali[4][i] == null) ? "" : ali[4][i].ToString() + " - ";
                    string DD = (ali[1][i] == null) ? " " : ali[1][i].ToString();

                    ins[i, 0] = DD;
                    ins[i, 2] = ali[0][i];
                    ins[i, 1] = ali[2][i];
                    ins[i, 3] = ali[3][i];
                    ins[i, 4] = ali[5][i];
                    ins[i, 5] = ali[1][i];
                }

                string resPath = new OpenExcelFile().Filenamereturn();
                if (resPath == "can not open file")
                    return;
                insert(resPath, ins, N, 5);
            }

            
            //МЕГА
            if (Combobox.Text == "Megafon")
            {
                object[][] ali = getarray(ComisPath,1, new int[] { 12, 17, 51, 59, 69, 60 }); // колонки меги

                int N = ali[0].Length;
                object[,] ins = new object[N, 6];

                for (int i = 0; i < N; i++)
                {
                    string TT = "";
                    if (ali[4][i] != null)
                        TT = ( ali[4][i].ToString() == "" ) ? "" : ali[4][i].ToString() + " - ";
                    string DD = (ali[3][i] == null) ? " " : ali[3][i].ToString();

                    ins[i, 0] = DD;
                    ins[i, 1] = dateinper(ali[1][i].ToString());
                    ins[i, 2] = (ali[2][i] == null) ? 0 : ali[2][i];
                    ins[i, 3] = ali[0][i];
                    ins[i, 4] = ali[5][i];
                    ins[i, 5] = ali[1][i];
                }

                string resPath = new OpenExcelFile().Filenamereturn();
                if (resPath == "can not open file")
                    return;
                insert(resPath, ins, N, 6);
            }

            win.MessageBox.Show("Конец");
        }

        static private double Predel(double nach, string tariff, int period)
        {
            if (nach > 6000)
                return 6000;

            return nach;
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
                            if (d.name.ToString().Contains(DD))
                                if (a == "a")
                                    d.inArhiv++;
                                else if (a == "b")
                                    d.inBase++;
                        }
                        break;
                    }
                }
            }
        }


    }
}
