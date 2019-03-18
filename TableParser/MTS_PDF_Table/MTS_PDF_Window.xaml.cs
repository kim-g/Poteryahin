using iTextSharp.text.pdf;
using System;
using System.Collections;
using System.Collections.Generic;
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

namespace MTS_PDF_Table
{
    /// <summary>
    /// Логика взаимодействия для MTS_PDF_Window.xaml
    /// </summary>
    public partial class MTS_PDF_Window : Window
    {
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

        private void FillForm()
        {
            string pdfTemplate = @"Standart.pdf";
            string newFile = @"Standart_Out.pdf";
            PdfReader pdfReader = new PdfReader(pdfTemplate);
            PdfStamper pdfStamper = new PdfStamper(pdfReader, new FileStream(
                newFile, FileMode.Create));
            AcroFields pdfFormFields = pdfStamper.AcroFields;
            // set form pdfFormFields
            // The first worksheet and W-4 form

            pdfFormFields.SetField("01", "01");
            pdfFormFields.SetField("02", "02");
            pdfFormFields.SetField("03", "03");
            pdfFormFields.SetField("04", "04");
            pdfFormFields.SetField("05", "05");
            pdfFormFields.SetField("06", "06");
            pdfFormFields.SetField("07", "07");
            pdfFormFields.SetField("08", "08");
            pdfFormFields.SetField("09", "09");
            pdfFormFields.SetField("10", "10");
            pdfFormFields.SetField("11", "11");
            pdfFormFields.SetField("12", "12");
            pdfFormFields.SetField("13", "13");
            pdfFormFields.SetField("14", "14");
            pdfFormFields.SetField("15", "15");
            pdfFormFields.SetField("16", "16");
            pdfFormFields.SetField("17", "17");
            pdfFormFields.SetField("18", "18");
            pdfFormFields.SetField("19", "19");
            pdfFormFields.SetField("20", "20");
            pdfFormFields.SetField("21", "21");
            pdfFormFields.SetField("22", "22");
            pdfFormFields.SetField("23", "23");
            pdfFormFields.SetField("24", "24");
            pdfFormFields.SetField("25", "25");
            pdfFormFields.SetField("26", "26");
            pdfFormFields.SetField("27", "27");
            pdfFormFields.SetField("28", "28");
            pdfFormFields.SetField("29", "29");
            pdfFormFields.SetField("30", "30");
            pdfFormFields.SetField("31", "31");
            pdfFormFields.SetField("32", "32");
            pdfFormFields.SetField("33", "33");
            pdfFormFields.SetField("34", "34");
            pdfFormFields.SetField("35", "35");
            pdfFormFields.SetField("36", "36");
            pdfFormFields.SetField("37", "37");
            pdfFormFields.SetField("38", "38");
            pdfFormFields.SetField("39", "39");
            pdfFormFields.SetField("40", "40");
            pdfFormFields.SetField("41", "41");
            pdfFormFields.SetField("42", "42");
            pdfFormFields.SetField("43", "43");
            pdfFormFields.SetField("44", "44");
            pdfFormFields.SetField("45", "45");
            pdfFormFields.SetField("46", "46");
            pdfFormFields.SetField("47", "47");
            pdfFormFields.SetField("48", "48");
            pdfFormFields.SetField("49", "49");
            pdfFormFields.SetField("50", "50");
            pdfFormFields.SetField("51", "51");
            pdfFormFields.SetField("52", "52");
            pdfFormFields.SetField("53", "53");
            pdfFormFields.SetField("54", "54");
            pdfFormFields.SetField("55", "55");
            pdfFormFields.SetField("56", "56");
            pdfFormFields.SetField("57", "57");
            pdfFormFields.SetField("58", "58");
            pdfFormFields.SetField("59", "59");
            pdfFormFields.SetField("60", "60");
            pdfFormFields.SetField("61", "61");
            pdfFormFields.SetField("62", "62");
            pdfFormFields.SetField("63", "63");
            pdfFormFields.SetField("64", "64");
            pdfFormFields.SetField("65", "65");
            pdfFormFields.SetField("66", "66");
            pdfFormFields.SetField("67", "67");
            pdfFormFields.SetField("68", "68");
            pdfFormFields.SetField("69", "69");
            pdfFormFields.SetField("text1", "text1");
            pdfFormFields.SetField("text2", "text2");


            MessageBox.Show("Finished");
            // flatten the form to remove editting options, set it to false
            // to leave the form open to subsequent manual edits

            pdfStamper.FormFlattening = false;
            // close the pdf

            pdfStamper.Close();
        }


        private void FilterExists_Click(object sender, RoutedEventArgs e)
        {
            FillForm();
        }

        private void SaveOutFile_Click(object sender, RoutedEventArgs e)
        {

        }

        private void ExcludeFilterFile_Click(object sender, RoutedEventArgs e)
        {

        }

        private void Aborting_Click(object sender, RoutedEventArgs e)
        {

        }
    }
}
