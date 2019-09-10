using ClosedXML.Excel;
using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows;
using System.Windows.Threading;
using Extentions;

namespace ExcelFilter
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private bool Abort = false;
        int LastPercent = -1;
        const bool ShowDifBtn = true;
        const int Head = 1;

        public MainWindow()
        {
            InitializeComponent();
            FilterТщеExists.Visibility = ShowDifBtn
                ? Visibility.Visible
                : Visibility.Collapsed;
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
            NoFilter.Visibility = Visibility.Collapsed;
            Aborting.Visibility = Visibility.Visible;

            // Открытие исходных файлов и подготовка DataTable таблиц
            StatusBlock.Text = "Открытие файла данных";
            Wait();
            if (Abort) return;
            DataTable In = LoadIn(FromTB.Text);
            StatusBlock.Text = "Открытие файла фильтров";
            Wait();
            if (Abort) return;

            string Operation = ((FrameworkElement)sender).Tag.ToString();

            foreach (object FilterFile in FilterTB.Items)
            {
                DataTable Filter = LoadFilters(FilterFile.ToString());
                string FilterName = Path.GetFileName(FilterFile.ToString());
                StatusBlock.Text = $"{FilterName}: Подготовка выходного файла";
                Wait();
                if (Abort) return;
                DataTable Out = PrepareOut(In);
                Wait();
                if (Abort) return;

                // Обработка фильтра
                switch (Operation)
                {
                    case "intersection":
                        Out = FindOverlap(In, Filter, Out, FilterName);
                        break;

                    case "difference":
                        Out = FindDifferences(In, Filter, Out, FilterName);
                        break;

                    case "lack of filters":
                        In = FindDifferences(In, Filter, Out, FilterName);
                        break;
                }
                if (Out == null || In  == null) return;

                // Сохраняем данные, если не поиск по ВСЕМ фильтрам
                if (new string[] { "intersection", "difference" }.Contains(Operation))
                {
                    StatusBlock.Text = $"{FilterName}: Сохранение данных";
                    Wait();
                    if (Abort) return;
                    SaveToXMLS(Out, Path.Combine(OutTB.Text, Path.GetFileName(FilterFile.ToString())));
                    Wait();
                    if (Abort) return;
                }
                StatusBlock.Text = $"{FilterName}: Данные отфильтрованы";
            }

            // Сохраняем данные ГЛОБАЛЬНОГО поиска
            if (new string[] { "lack of filters" }.Contains(Operation))
            {
                string FilterName = "AbsenceInFilters.xlsx";
                StatusBlock.Text = $"{FilterName}: Сохранение данных";
                Wait();
                if (Abort) return;
                SaveToXMLS(In, Path.Combine(OutTB.Text, FilterName));
                Wait();
                if (Abort) return;
            }

            StatusBlock.Text = $"Выберите файлы и операцию";
            MessageBox.Show("Все данные отфильтрованы");
            FilterExists.Visibility = Visibility.Visible;
            FilterТщеExists.Visibility = ShowDifBtn
                ? Visibility.Visible
                : Visibility.Collapsed;
            NoFilter.Visibility = Visibility.Visible;
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
            IXLWorksheet CurWorksheet = OutTable.Worksheets.Add("Отфильтрованные данные");

            // Добавление заголовка
            int i = 1;
            foreach (DataColumn Colomn in Out.Columns)
                CurWorksheet.Cell(1, i++).Value = Colomn.Caption;

            // Добавление данных
            i = 0;
            string StatusStr = $"{Path.GetFileName(FileName)}: Сохранение данных";
            int m = Out.Rows.Count;
            SetStatus(StatusStr, i, m);
            foreach (DataRow R in Out.Rows)
            {
                int j = 0;
                foreach (DataColumn Colomn in Out.Columns)
                {
                    CurWorksheet.Cell(i + 1 + Head, j + 1).Value =
                        Out.Rows[i].ItemArray[j];
                    if (Out.Columns[j].DataType == typeof(DateTime))
                    {
                        CurWorksheet.Cell(i + 1 + Head, j + 1).Value =
                            Out.Rows[i].ItemArray[j];
                        CurWorksheet.Cell(i + 1 + Head, j + 1).DataType = XLDataType.DateTime;
                    }
                    else
                    {
                        CurWorksheet.Cell(i + 1 + Head, j + 1).SetValue(Out.Rows[i].ItemArray[j].ToString());
                    }
                    j++;
                }
                i++;
                SetStatus(StatusStr, i, m);
                Wait();
            }
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
            int k = 1;
            foreach (object X in FromSheet.Columns())
            {
                if (FromSheet.RangeUsed().RowsUsed().ToArray()[0].Cell(k).Value.ToString() == "") continue;

                In.Columns.Add(FromSheet.RangeUsed().RowsUsed().ToArray()[0].Cell(k).Value.ToString(),
                    FromSheet.RangeUsed().RowsUsed().ToArray()[1].Cell(k++).Value.GetType());
            }

            // Подготовка счётчиков для статусной строки
            string StatusStr = "Загрузка данных";
            int i = 0;
            int m = FromSheet.RangeUsed().RowsUsed().Skip(Head).Count();
            SetStatus(StatusStr, i, m);

            // Загрузка данных из книги Excel в DataTable
            foreach (var row in FromSheet.RangeUsed().RowsUsed().Skip(Head))
            {
                object[] NewRow = new object[In.Columns.Count];
                for (int j = 0; j < In.Columns.Count; j++)
                    NewRow[j] = row.Cell(j + 1).Value;
                In.Rows.Add(NewRow);
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
            XLWorkbook FilterTable = new XLWorkbook(FileName);
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
        private DataTable PrepareOut(DataTable In)
        {
            DataTable Out = new DataTable();
            Out.TableName = "Отфильтрованные данные"; 
            foreach (DataColumn X in In.Columns)
                Out.Columns.Add(X.Caption, X.DataType);

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
                StatusBlock.Text = StatusString + ": " + NewPercent.ToString() + "%";
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
                FilterТщеExists.Visibility = ShowDifBtn
                    ? Visibility.Visible
                    : Visibility.Collapsed;
                NoFilter.Visibility = Visibility.Visible;
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
                foreach (object InCell in inrow.ItemArray)
                {
                    bool Stop = false;
                    foreach (DataRow frow in Filter.Rows)
                    {
                        if (InCell.ToString() == frow.ItemArray[0].ToString())
                        {
                            Out.Rows.Add(inrow.ItemArray);
                            Stop = true;
                            break;
                        }
                    }
                    if (Stop) break;
                }
                SetStatus(StatusStr, i++, m);
                Wait();
                if (Abort) return null;
            }

            return Out;
        }

        /// <summary>
        /// Ищет разницу исходных данных с фильтром
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
                bool Stop = false;
                foreach (object InCell in inrow.ItemArray)
                {
                    foreach (DataRow frow in Filter.Rows)
                    {
                        if (InCell.ToString() == frow.ItemArray[0].ToString())
                        {
                            Stop = true;
                            break;
                        }
                        if (Stop) break;
                    }
                    if (Stop) break;
                }
                if (!Stop) Out.Rows.Add(inrow.ItemArray);
                SetStatus(StatusStr, i++, m);
                Wait();
                if (Abort) return null;
            }

            return Out;
        }
    }
}
