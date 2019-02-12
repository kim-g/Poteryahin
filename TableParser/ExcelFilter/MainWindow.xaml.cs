using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading;
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
using System.Windows.Threading;

namespace ExcelFilter
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private bool Abort = false;
        int LastPercent = -1;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void OpenFromFile_Click(object sender, RoutedEventArgs e)
        {
            string Answer = Files.OpenFile("Открыть файл с исходными данными");
            if (Answer != null)
                FromTB.Text = Answer;
        }

        private void SaveOutFile_Click(object sender, RoutedEventArgs e)
        {
            string Answer = Files.SaveFile("Открыть файл с исходными данными");
            if (Answer != null)
                OutTB.Text = Answer;
        }

        private void FilterExists_Click(object sender, RoutedEventArgs e)
        {
            // Установка возможности прерывания
            Abort = false;
            FilterExists.Visibility = Visibility.Collapsed;
            FilterТщеExists.Visibility = Visibility.Collapsed;
            Aborting.Visibility = Visibility.Visible;

            string StatusStr;
            Status.Content = "Открытие файла данных";
            Wait();
            if (Abort) return;
            DataTable In = LoadIn(FromTB.Text);
            Status.Content = "Открытие файла фильтров";
            Wait();
            if (Abort) return;
            DataTable Filter = LoadFilters(FilterTB.Text);
            Status.Content = "Подготовка выходного файла";
            Wait();
            if (Abort) return;
            DataTable Out = PrepareOut();
            Wait();
            if (Abort) return;

            StatusStr = "Поиск совпадений";
            int i = 0;
            int m = In.Rows.Count;
            SetStatus(StatusStr, i, m);
            Wait();
            if (Abort) return;
            // Ищем совпадения
            foreach (DataRow inrow in In.Rows)
            {
                foreach (DataRow frow in Filter.Rows)
                {
                    if (inrow.ItemArray[0].ToString() == frow.ItemArray[0].ToString())
                    {
                        Out.Rows.Add(inrow.ItemArray);
                        break;
                    }
                }
                SetStatus(StatusStr, i++, m);
                Wait();
                if (Abort) return;
            }

            Status.Content = "Сохранение данных";
            Wait();
            if (Abort) return;
            SaveToXMLS(Out);
            Wait();
            if (Abort) return;
            Status.Content = "Данные отфильтрованы";
            MessageBox.Show("Данные отфильтрованы");
            FilterExists.Visibility = Visibility.Visible;
            FilterТщеExists.Visibility = Visibility.Visible;
            Aborting.Visibility = Visibility.Collapsed;
        }

        private void SaveToXMLS(DataTable Out)
        {
            XLWorkbook OutTable = new XLWorkbook();
            Wait();
            OutTable.Worksheets.Add(Out);
            Wait();
            OutTable.SaveAs(OutTB.Text);
        }

        private DataTable LoadIn(string FileName)
        {
            XLWorkbook FromTable = new XLWorkbook(FileName);
            IXLWorksheet FromSheet = FromTable.Worksheets.ToList()[0];

            DataTable In = new DataTable();
            In.TableName = "Исходные данные";
            In.Columns.Add("Number", typeof(string));
            In.Columns.Add("Data1", typeof(string));
            In.Columns.Add("Data2", typeof(string));
            In.Columns.Add("Date", typeof(DateTime));

            string StatusStr = "Загрузка данных";
            int i = 0;
            int m = FromSheet.RangeUsed().RowsUsed().Skip(0).Count();
            SetStatus(StatusStr, i, m);
            foreach (var row in FromSheet.RangeUsed().RowsUsed().Skip(0))
            {
                In.Rows.Add(row.Cell(1).Value, row.Cell(2).Value, row.Cell(3).Value, row.Cell(4).Value);
                SetStatus(StatusStr, i++, m);
                Wait();
            }

            return In;
        }

        private DataTable LoadFilters(string FileName)
        {
            XLWorkbook FilterTable = new XLWorkbook(FilterTB.Text);
            IXLWorksheet FilterSheet = FilterTable.Worksheets.ToList()[0];

            DataTable Filter = new DataTable();
            Filter.TableName = "Фильтры";
            Filter.Columns.Add("Number", typeof(string));

            string StatusStr = "Загрузка фильтров";
            int i = 0;
            int m = FilterSheet.RangeUsed().RowsUsed().Skip(0).Count();
            SetStatus(StatusStr, i, m);

            foreach (var row in FilterSheet.RangeUsed().RowsUsed().Skip(0))
            {
                Filter.Rows.Add(row.Cell(1).Value);
                SetStatus(StatusStr, i++, m);
                Wait();
            }

            return Filter;
        }

        private DataTable PrepareOut()
        {
            DataTable Out = new DataTable();
            Out.TableName = "Отфильтрованные данные";
            Out.Columns.Add("Number", typeof(string));
            Out.Columns.Add("Data1", typeof(string));
            Out.Columns.Add("Data2", typeof(string));
            Out.Columns.Add("Date", typeof(DateTime));

            return Out;
        }

        private void OpenFilterFile_Click(object sender, RoutedEventArgs e)
        {
            string Answer = Files.OpenFile("Открыть файл с исходными данными");
            if (Answer != null)
                FilterTB.Text = Answer;
        }

        private void Wait()
        {
            Application.Current.Dispatcher.Invoke(DispatcherPriority.Background, new ThreadStart(delegate { }));
        }

        private void SetStatus(string StatusString, int Pos, int Max)
        {
            int NewPercent = Pos * 100 / Max;
            if (NewPercent != LastPercent)
            {
                Status.Content = StatusString + ": " + NewPercent.ToString() + "%";
                LastPercent = NewPercent;
            }
        }

        private void FilterТщеExists_Click(object sender, RoutedEventArgs e)
        {
            // Установка возможности прерывания
            Abort = false;
            FilterExists.Visibility = Visibility.Collapsed;
            FilterТщеExists.Visibility = Visibility.Collapsed;
            Aborting.Visibility = Visibility.Visible;

            string StatusStr;
            Status.Content = "Открытие файла данных";
            Wait();
            if (Abort) return;
            DataTable In = LoadIn(FromTB.Text);
            Status.Content = "Открытие файла фильтров";
            Wait();
            if (Abort) return;
            DataTable Filter = LoadFilters(FilterTB.Text);
            Status.Content = "Подготовка выходного файла";
            Wait();
            if (Abort) return;
            DataTable Out = PrepareOut();
            Wait();
            if (Abort) return;

            // Ищем совпадения
            int i = 0;
            int m = In.Rows.Count;
            StatusStr = "Поиск несовпадающих записей";
            SetStatus(StatusStr, i, m);
            Wait();
            if (Abort) return;
            foreach (DataRow inrow in In.Rows)
            {
                bool Found = false;
                foreach (DataRow frow in Filter.Rows)
                {
                    if (inrow.ItemArray[0].ToString() == frow.ItemArray[0].ToString())
                    {
                        Found = true;
                        break;
                    }
                }
                SetStatus(StatusStr, i++, m);
                if (!Found) Out.Rows.Add(inrow.ItemArray);
                Wait();
                if (Abort) return;
            }

            Status.Content = "Сохранение данных";
            Wait();
            if (Abort) return;
            SaveToXMLS(Out);
            Status.Content = "Данные отфильтрованы";
            MessageBox.Show("Данные отфильтрованы");
            FilterExists.Visibility = Visibility.Visible;
            FilterТщеExists.Visibility = Visibility.Visible;
            Aborting.Visibility = Visibility.Collapsed;
        }

        private void Aborting_Click(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show("Вы уверены, что хотите прервать процесс?", "Прерывание", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
            {
                Abort = true;
                FilterExists.Visibility = Visibility.Visible;
                FilterТщеExists.Visibility = Visibility.Visible;
                Aborting.Visibility = Visibility.Collapsed;
            }
        }
    }
}
