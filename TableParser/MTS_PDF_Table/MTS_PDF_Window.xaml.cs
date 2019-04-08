using ClosedXML.Excel;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows;

using System.Windows.Threading;

namespace MTS_PDF_Table
{
    /// <summary>
    /// Логика взаимодействия для MTS_PDF_Window.xaml
    /// </summary>
    public partial class MTS_PDF_Window : Window
    {
        const int Head = 1;
        bool Abort = false;
        int LastPercent = 143;
        public static Log LogWindow = new Log();

        public MTS_PDF_Window()
        {
            InitializeComponent();
        }

        private void OpenFilterFile_Click(object sender, RoutedEventArgs e)
        {
            string[] Answer = ExcelFilter.Files.OpenFiles("Открыть файл с исходными данными");
            if (Answer != null)
                foreach (string X in Answer)
                    FilterTB.Items.Add(X);
        }

        /// <span class="code-SummaryComment"><summary></span>
        /// List all of the form fields into a textbox. The
        /// form fields identified can be used to map each of the
        /// fields in a PDF.
        /// <span class="code-SummaryComment"></summary></span>
        private void ListFieldNames()
        {
            string pdfTemplate = @"Standart.pdf";
            // title the form

            this.Title += " - " + pdfTemplate;
            // create a new PDF reader based on the PDF template document

            PdfReader pdfReader = new PdfReader(pdfTemplate);
            // create and populate a string builder with each of the
            // field names available in the subject PDF

            StringBuilder sb = new StringBuilder();
            foreach (KeyValuePair<string,AcroFields.Item> de in pdfReader.AcroFields.Fields)
            {
                FilterTB.Items.Add(de.Key.ToString());
            }
        }

        private void FilterExists_Click(object sender, RoutedEventArgs e)
        {
            foreach (string InFile in FilterTB.Items)
            {
                string PureName = Path.GetFileNameWithoutExtension(InFile);
                DataTable InTable = LoadIn(InFile);

                // Подготовка счётчиков для статусной строки
                string StatusStr = $"{PureName}: Заполнение форм";
                int m = InTable.Rows.Count;
                SetStatus(StatusStr, 0, m);

                for (int i = 0; i < InTable.Rows.Count; i++)
                {
                    PersonInfo PI = new PersonInfo(InTable, i);
                    PI.FillForm(Path.Combine(OutTB.Text, $"{PI.Line} - {PI.Number}.pdf"));
                    SetStatus(StatusStr, i, m);
                    Wait();
                    if (Abort) return;
                }
            }

            StatusBlock.Text = "Обработка форм завершена.";
        }

        private void SaveOutFile_Click(object sender, RoutedEventArgs e)
        {
            string Answer = ExcelFilter.Files.OpenDirectory("Открыть файл с исходными данными");
            if (Answer != null)
                OutTB.Text = Answer;
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

        private void Aborting_Click(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show("Вы уверены, что хотите прервать процесс?", "Прерывание", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
            {
                Abort = true;
                FilterExists.Visibility = Visibility.Visible;
                Aborting.Visibility = Visibility.Collapsed;
            }
        }

        /// <summary>
        /// Загрузка исходных данных в таблицу DataTable
        /// </summary>
        /// <param name="FileName">Имя файла, из которого загружаются данные</param>
        /// <returns></returns>
        private DataTable LoadIn(string FileName)
        {
            // Загрузка книги Excel
            string PureName = Path.GetFileNameWithoutExtension(FileName);
            XLWorkbook FromTable = new XLWorkbook(FileName);
            IXLWorksheet FromSheet = FromTable.Worksheets.ToList()[0];

            // Подготовка таблицы
            DataTable In = new DataTable();
            In.TableName = PureName;
            int k = 1;
            foreach (object X in FromSheet.Columns())
                In.Columns.Add(FromSheet.RangeUsed().RowsUsed().ToArray()[0].Cell(k).Value.ToString(),
                    FromSheet.RangeUsed().RowsUsed().ToArray()[1].Cell(k++).Value.GetType());

            // Подготовка счётчиков для статусной строки
            string StatusStr = $"{PureName}: Загрузка данных";
            int i = 0;
            int m = FromSheet.RangeUsed().RowsUsed().Skip(Head).Count();
            SetStatus(StatusStr, i, m);
            Wait();

            // Загрузка данных из книги Excel в DataTable
            foreach (var row in FromSheet.RangeUsed().RowsUsed().Skip(Head))
            {
                object[] NewRow = new object[In.Columns.Count];
                for (int j = 0; j < In.Columns.Count; j++)
                    NewRow[j] = row.Cell(j + 1).Value;
                try
                {
                    In.Rows.Add(NewRow);
                }
                catch (Exception e)
                {
                    LogWindow.Add($"{PureName}: строка {row.RowNumber()}: ошибка загрузки данных (неправильный формат): {e.Message}");
                }
                SetStatus(StatusStr, i++, m);
                Wait();
            }

            return In;
        }

        /// <summary>
        /// Прерывает выполнение процесса для обработки поступивших событий
        /// </summary>
        private void Wait()
        {
            Application.Current.Dispatcher.Invoke(new ThreadStart(delegate { }), DispatcherPriority.Background);
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

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            LogWindow.Close();
        }
    }
}
