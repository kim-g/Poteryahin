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
            Status.Content = "Загрузка данных";
            DataTable In = LoadIn(FromTB.Text);
            DataTable Filter = LoadFilters(FilterTB.Text);
            DataTable Out = PrepareOut();

            Status.Content = "Поиск совпадений";
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
            }

            Status.Content = "Сохранение данных";
            SaveToXMLS(Out);
            Status.Content = "Данные отфильтрованы";
            MessageBox.Show("Данные отфильтрованы");
        }

        private void SaveToXMLS(DataTable Out)
        {
            XLWorkbook OutTable = new XLWorkbook();
            OutTable.Worksheets.Add(Out);

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

            foreach (var row in FromSheet.RangeUsed().RowsUsed().Skip(0))
                In.Rows.Add(row.Cell(1).Value, row.Cell(2).Value, row.Cell(3).Value, row.Cell(4).Value);

            return In;
        }

        private DataTable LoadFilters(string FileName)
        {
            XLWorkbook FilterTable = new XLWorkbook(FilterTB.Text);
            IXLWorksheet FilterSheet = FilterTable.Worksheets.ToList()[0];

            DataTable Filter = new DataTable();
            Filter.TableName = "Фильтры";
            Filter.Columns.Add("Number", typeof(string));

            foreach (var row in FilterSheet.RangeUsed().RowsUsed().Skip(0))
                Filter.Rows.Add(row.Cell(1).Value);

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

        private void FilterТщеExists_Click(object sender, RoutedEventArgs e)
        {
            Status.Content = "Загрузка данных";
            Wait();
            DataTable In = LoadIn(FromTB.Text);
            DataTable Filter = LoadFilters(FilterTB.Text);
            DataTable Out = PrepareOut();

            // Ищем совпадения
            Status.Content = "Поиск несовпадающих записей";
            Wait();
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
                if (!Found) Out.Rows.Add(inrow.ItemArray);
            }

            Status.Content = "Сохранение данных";
            Wait();
            SaveToXMLS(Out);
            Status.Content = "Данные отфильтрованы";
            MessageBox.Show("Данные отфильтрованы");
        }
    }
}
