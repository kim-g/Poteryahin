using ClosedXML.Excel;
using iTextSharp.text.pdf;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
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

namespace MTS_PDF_Table
{
    /// <summary>
    /// Логика взаимодействия для MTS_PDF_Window.xaml
    /// </summary>
    public partial class MTS_PDF_Window : Window
    {
        const int Head = 1;
        bool Abort = false;

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
                DataTable InTable = LoadIn(InFile);
                PersonInfo PI = new PersonInfo(InTable, 0);
                PI.FillForm(@"Standart_Out.pdf");
            }
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
            XLWorkbook FromTable = new XLWorkbook(FileName);
            IXLWorksheet FromSheet = FromTable.Worksheets.ToList()[0];

            // Подготовка таблицы
            DataTable In = new DataTable();
            In.TableName = "Исходные данные";
            int k = 1;
            foreach (object X in FromSheet.Columns())
                In.Columns.Add(FromSheet.RangeUsed().RowsUsed().ToArray()[0].Cell(k).Value.ToString(),
                    FromSheet.RangeUsed().RowsUsed().ToArray()[1].Cell(k++).Value.GetType());

            // Подготовка счётчиков для статусной строки
            string StatusStr = "Загрузка данных";
            int i = 0;
            int m = FromSheet.RangeUsed().RowsUsed().Skip(Head).Count();
            //SetStatus(StatusStr, i, m);

            // Загрузка данных из книги Excel в DataTable
            foreach (var row in FromSheet.RangeUsed().RowsUsed().Skip(Head))
            {
                object[] NewRow = new object[In.Columns.Count];
                for (int j = 0; j < In.Columns.Count; j++)
                    NewRow[j] = row.Cell(j + 1).Value;
                In.Rows.Add(NewRow);
                //SetStatus(StatusStr, i++, m);
                Wait();
            }

            return In;
        }

        /// <summary>
        /// Прерывает выполнение процесса для обработки поступивших событий
        /// </summary>
        private void Wait()
        {
            Application.Current.Dispatcher.Invoke(DispatcherPriority.Background, new ThreadStart(delegate { }));
        }
    }
}
