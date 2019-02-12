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

        /// <summary>
        /// Открывает диалоговое окно открытия файла данных и переносит путь файла в текстовое поле
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OpenFromFile_Click(object sender, RoutedEventArgs e)
        {
            string Answer = Files.OpenFile("Открыть файл с исходными данными");
            if (Answer != null)
                FromTB.Text = Answer;
        }

        /// <summary>
        /// Открывает диалоговое окно сохранения файла отфильтрованных данных и переносит путь файла в текстовое поле
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SaveOutFile_Click(object sender, RoutedEventArgs e)
        {
            string Answer = Files.SaveFile("Открыть файл с исходными данными");
            if (Answer != null)
                OutTB.Text = Answer;
        }

        /// <summary>
        /// Фильтрует данные по совпадению
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void FilterExists_Click(object sender, RoutedEventArgs e)
        {
            // Установка возможности прерывания
            Abort = false;
            FilterExists.Visibility = Visibility.Collapsed;
            FilterТщеExists.Visibility = Visibility.Collapsed;
            Aborting.Visibility = Visibility.Visible;

            // Открытие исходных файлов и подготовка DataTable таблиц
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

            // Подготовка счётчика для статусной строки
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

            // Сохраняем данные
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

        /// <summary>
        /// Сохранение данных в xlsx формате. Имя файла берётся из текстового поля
        /// </summary>
        /// <param name="Out">Таблица для сохранения</param>
        private void SaveToXMLS(DataTable Out)
        {
            XLWorkbook OutTable = new XLWorkbook();
            Wait();
            OutTable.Worksheets.Add(Out);
            Wait();
            OutTable.SaveAs(OutTB.Text);
        }

        /// <summary>
        /// Загрузка исходных данных в таблицу DataTable
        /// </summary>
        /// <param name="FileName">Имя файла, из которого загружаются данные</param>
        /// <returns></returns>
        private DataTable LoadIn(string FileName)
        {
            // Загрузка книги Excel
            XLWorkbook FromTable = new XLWorkbook(FileName);
            IXLWorksheet FromSheet = FromTable.Worksheets.ToList()[0];

            // Подготовка таблицы
            DataTable In = new DataTable();
            In.TableName = "Исходные данные";
            In.Columns.Add("Number", typeof(string));
            In.Columns.Add("Data1", typeof(string));
            In.Columns.Add("Data2", typeof(string));
            In.Columns.Add("Date", typeof(DateTime));

            // Подготовка счётчиков для статусной строки
            string StatusStr = "Загрузка данных";
            int i = 0;
            int m = FromSheet.RangeUsed().RowsUsed().Skip(0).Count();
            SetStatus(StatusStr, i, m);

            // Загрузка данных из книги Excel в DataTable
            foreach (var row in FromSheet.RangeUsed().RowsUsed().Skip(0))
            {
                In.Rows.Add(row.Cell(1).Value, row.Cell(2).Value, row.Cell(3).Value, row.Cell(4).Value);
                SetStatus(StatusStr, i++, m);
                Wait();
            }

            return In;
        }

        /// <summary>
        /// Загрузка фильтров в таблицу DataTable
        /// </summary>
        /// <param name="FileName">Имя файла, из которого загружаются фильтры</param>
        /// <returns></returns>
        private DataTable LoadFilters(string FileName)
        {
            // Загрузка книги Excel
            XLWorkbook FilterTable = new XLWorkbook(FilterTB.Text);
            IXLWorksheet FilterSheet = FilterTable.Worksheets.ToList()[0];

            // Подготовка таблицы
            DataTable Filter = new DataTable();
            Filter.TableName = "Фильтры";
            Filter.Columns.Add("Number", typeof(string));

            // Подготовка счётчиков для статусной строки
            string StatusStr = "Загрузка фильтров";
            int i = 0;
            int m = FilterSheet.RangeUsed().RowsUsed().Skip(0).Count();
            SetStatus(StatusStr, i, m);

            // Загрузка фильтров из книги Excel в DataTable
            foreach (var row in FilterSheet.RangeUsed().RowsUsed().Skip(0))
            {
                Filter.Rows.Add(row.Cell(1).Value);
                SetStatus(StatusStr, i++, m);
                Wait();
            }

            return Filter;
        }

        /// <summary>
        /// Подготовка выводной таблицы
        /// </summary>
        /// <returns></returns>
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

        /// <summary>
        /// Открывает диалоговое окно открытия файла фильтров и переносит путь файла в текстовое поле
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OpenFilterFile_Click(object sender, RoutedEventArgs e)
        {
            string Answer = Files.OpenFile("Открыть файл с исходными данными");
            if (Answer != null)
                FilterTB.Text = Answer;
        }

        /// <summary>
        /// Прерывает выполнение процесса для обработки поступивших событий
        /// </summary>
        private void Wait()
        {
            Application.Current.Dispatcher.Invoke(DispatcherPriority.Background, new ThreadStart(delegate { }));
        }

        /// <summary>
        /// Устанавливает статусую строку с процентом выполнения. 
        /// </summary>
        /// <param name="StatusString">Текст статусной строки</param>
        /// <param name="Pos">Позиция выполнения</param>
        /// <param name="Max">Максимальная позиция выполнения</param>
        private void SetStatus(string StatusString, int Pos, int Max)
        {
            // Меняет статус только в том случае, если изменяется процент выполнения для ускорения работы.
            int NewPercent = Pos * 100 / Max;
            if (NewPercent != LastPercent)
            {
                Status.Content = StatusString + ": " + NewPercent.ToString() + "%";
                LastPercent = NewPercent;
            }
        }
        /// <summary>
        /// Фильтрует данные по несовпадению
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void FilterТщеExists_Click(object sender, RoutedEventArgs e)
        {
            // Установка возможности прерывания
            Abort = false;
            FilterExists.Visibility = Visibility.Collapsed;
            FilterТщеExists.Visibility = Visibility.Collapsed;
            Aborting.Visibility = Visibility.Visible;

            // Загрузка исходных данных
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

            // Подготовка счётчика статусной строки
            int i = 0;
            int m = In.Rows.Count;
            StatusStr = "Поиск несовпадающих записей";
            SetStatus(StatusStr, i, m);
            Wait();
            if (Abort) return;
            // Ищем совпадения
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

            // Сохранение отфильтрованных данных
            Status.Content = "Сохранение данных";
            Wait();
            if (Abort) return;
            SaveToXMLS(Out);
            Status.Content = "Данные отфильтрованы";
            MessageBox.Show("Данные отфильтрованы");

            // Возвращаем кнопки на место
            FilterExists.Visibility = Visibility.Visible;
            FilterТщеExists.Visibility = Visibility.Visible;
            Aborting.Visibility = Visibility.Collapsed;
        }

        /// <summary>
        /// Нажатие на кнопку прерывания работы
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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
