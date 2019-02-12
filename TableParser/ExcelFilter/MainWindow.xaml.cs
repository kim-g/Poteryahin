using ClosedXML.Excel;
using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows;
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
            string Answer = Files.OpenDirectory("Открыть файл с исходными данными");
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
            Status.Content = "Открытие файла данных";
            Wait();
            if (Abort) return;
            DataTable In = LoadIn(FromTB.Text);
            Status.Content = "Открытие файла фильтров";
            Wait();
            if (Abort) return;

            foreach (object FilterFile in FilterTB.Items)
            {
                DataTable Filter = LoadFilters(FilterFile.ToString());
                string FilterName = Path.GetFileName(FilterFile.ToString());
                Status.Content = $"{FilterName}: Подготовка выходного файла";
                Wait();
                if (Abort) return;
                DataTable Out = PrepareOut();
                Wait();
                if (Abort) return;

                // Обработка фильтра
                Out = ((FrameworkElement)sender).Tag.ToString() == "0"
                    ? FindOverlap(In, Filter, Out, FilterName)
                    : FindDifferences(In, Filter, Out, FilterName);
                if (Out == null) return;

                // Сохраняем данные
                Status.Content = $"{FilterName}: Сохранение данных";
                Wait();
                if (Abort) return;
                SaveToXMLS(Out, Path.Combine(OutTB.Text, Path.GetFileName(FilterFile.ToString())));
                Wait();
                if (Abort) return;
                Status.Content = $"{FilterName}: Данные отфильтрованы";
            }
            MessageBox.Show("Все данные отфильтрованы");
            FilterExists.Visibility = Visibility.Visible;
            FilterТщеExists.Visibility = Visibility.Visible;
            Aborting.Visibility = Visibility.Collapsed;
        }

        /// <summary>
        /// Сохранение данных в xlsx формате. Имя файла берётся из текстового поля
        /// </summary>
        /// <param name="Out">Таблица для сохранения</param>
        private void SaveToXMLS(DataTable Out, string FileName)
        {
            XLWorkbook OutTable = new XLWorkbook();
            Wait();
            OutTable.Worksheets.Add(Out);
            Wait();
            OutTable.SaveAs(FileName);
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
            XLWorkbook FilterTable = new XLWorkbook(FilterTB.Items[0].ToString());
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
            string[] Answer = Files.OpenFiles("Открыть файл с исходными данными");
            if (Answer != null)
                foreach (string X in Answer)
                    FilterTB.Items.Add(X);
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

        private void ExcludeFilterFile_Click(object sender, RoutedEventArgs e)
        {
            if (FilterTB.SelectedItems.Count > 0)
            {
                object[] Selected = new object[FilterTB.SelectedItems.Count];
                FilterTB.SelectedItems.CopyTo(Selected, 0);
                foreach (object X in Selected)
                    FilterTB.Items.Remove(X);
            }
        }

        /// <summary>
        /// Ищет пересечения исходных данных с фильтром
        /// </summary>
        /// <param name="In">Входные данные</param>
        /// <param name="Filter">Фильтр</param>
        /// <param name="Out">Выходной формат</param>
        /// <returns></returns>
        private DataTable FindOverlap(DataTable In, DataTable Filter, DataTable Out, string FilterName)
        {
            string StatusStr = $"{FilterName}: Поиск совпадений";
            int i = 0;
            int m = In.Rows.Count;
            SetStatus(StatusStr, i, m);
            Wait();
            if (Abort) return null;
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
                if (Abort) return null;
            }

            return Out;
        }

        /// <summary>
        /// Ищет hfpybwe исходных данных с фильтром
        /// </summary>
        /// <param name="In">Входные данные</param>
        /// <param name="Filter">Фильтр</param>
        /// <param name="Out">Выходной формат</param>
        /// <returns></returns>
        private DataTable FindDifferences(DataTable In, DataTable Filter, DataTable Out, string FilterName)
        {
            string StatusStr = "Поиск совпадений";
            int i = 0;
            int m = In.Rows.Count;
            StatusStr = StatusStr = $"{FilterName}: Поиск несовпадающих записей";
            SetStatus(StatusStr, i, m);
            Wait();
            if (Abort) return null;
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
                if (Abort) return null;
            }

            return Out;
        }
    }
}
