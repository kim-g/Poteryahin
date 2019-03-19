using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace MTS_PDF_Table
{
    class PersonInfo
    {
        public string Number { get; set; }
        public string ICC { get; set; }
        public string ICC_Suffix { get; set; }
        public string Surname { get; set; }
        public string Name { get; set; }
        public string SecondName { get; set; }
        public bool Sex { get; set; }
        public string Document { get; set; }
        public string DocumentSerie { get; set; }
        public string DocumentNumber { get; set; }
        public string DocumentIssuedBy { get; set; }
        public DateTime DocumentIssueDate { get; set; }
        public DateTime Birth { get; set; }
        public string PlaceIndex { get; set; }
        public string PlaceCity { get; set; }
        public string PlaceStreet { get; set; }
        public string PlaceBuilding { get; set; }
        public string PlaceFlat { get; set; }
        public string Rate { get; set; }
        public DateTime ContractConclusionDate { get; set; }
        public string ContractConclusionPlace { get; set; }
        public string Seller { get; set; }
        public string SellerID { get; set; }

        const string pdfTemplate = @"Standart.pdf";

        /// <summary>
        /// Создаёт пустой объект PersonInfo 
        /// </summary>
        public PersonInfo()
        {

        }

        /// <summary>
        /// Создаёт объект PersonInfo из строки таблицы данных
        /// </summary>
        public PersonInfo(DataTable InTable, int Position)
        {
            foreach (DataColumn Col in InTable.Columns)
            {
                switch (Col.Caption.ToLowerInvariant())
                {
                    case "номер":
                        Number = InTable.Rows[Position].ItemArray[Col.Ordinal].ToString();
                        break;
                    case "icc":
                        string[] ICC_Arr = InTable.Rows[Position].ItemArray[Col.Ordinal].ToString().Split(' ');
                        ICC = ICC_Arr[0];
                        ICC_Suffix = ICC_Arr[1];
                        break;
                    case "фамилия":
                        Surname = InTable.Rows[Position].ItemArray[Col.Ordinal].ToString();
                        break;
                    case "имя":
                        Name = InTable.Rows[Position].ItemArray[Col.Ordinal].ToString();
                        break;
                    case "отчество":
                        SecondName = InTable.Rows[Position].ItemArray[Col.Ordinal].ToString();
                        break;
                    case "пол":
                        Sex = InTable.Rows[Position].ItemArray[Col.Ordinal].ToString().ToLowerInvariant() == "м";
                        break;
                    case "документ удост личность":
                        Document = InTable.Rows[Position].ItemArray[Col.Ordinal].ToString();
                        break;
                    case "серия":
                        DocumentSerie = InTable.Rows[Position].ItemArray[Col.Ordinal].ToString();
                        break;
                    case "номер документа":
                        DocumentNumber = InTable.Rows[Position].ItemArray[Col.Ordinal].ToString();
                        break;
                    case "кем выдан":
                        DocumentIssuedBy = InTable.Rows[Position].ItemArray[Col.Ordinal].ToString();
                        break;
                    case "дата выдачи":
                        DocumentIssueDate = DateTime.Parse(InTable.Rows[Position].ItemArray[Col.Ordinal].ToString());
                        break;
                    case "дата рождения":
                        Birth = DateTime.Parse(InTable.Rows[Position].ItemArray[Col.Ordinal].ToString());
                        break;
                    case "индекс":
                        PlaceIndex = InTable.Rows[Position].ItemArray[Col.Ordinal].ToString();
                        break;
                    case "город":
                        PlaceCity = InTable.Rows[Position].ItemArray[Col.Ordinal].ToString();
                        break;
                    case "улица":
                        PlaceStreet = InTable.Rows[Position].ItemArray[Col.Ordinal].ToString();
                        break;
                    case "дом":
                        PlaceBuilding = InTable.Rows[Position].ItemArray[Col.Ordinal].ToString();
                        break;
                    case "кв.":
                        PlaceFlat = InTable.Rows[Position].ItemArray[Col.Ordinal].ToString();
                        break;
                    case "тариф":
                        Rate = InTable.Rows[Position].ItemArray[Col.Ordinal].ToString();
                        break;
                    case "дата заключения догвора":
                        ContractConclusionDate = DateTime.Parse(InTable.Rows[Position].ItemArray[Col.Ordinal].ToString());
                        break;
                    case "место заключения":
                        ContractConclusionPlace = InTable.Rows[Position].ItemArray[Col.Ordinal].ToString();
                        break;
                    case "продавец":
                        Seller = InTable.Rows[Position].ItemArray[Col.Ordinal].ToString();
                        break;
                    case "код торговой точки":
                        SellerID = InTable.Rows[Position].ItemArray[Col.Ordinal].ToString();
                        break;
                }
            }
        }

        public void FillForm(string newFile)
        {
            PdfReader pdfReader = new PdfReader(pdfTemplate);
            PdfStamper pdfStamper = new PdfStamper(pdfReader, new FileStream(
                newFile, FileMode.Create));
            AcroFields pdfFormFields = pdfStamper.AcroFields;
            // set form pdfFormFields
            // The first worksheet and W-4 form

            /*String[] values = pdfFormFields.GetAppearanceStates("10");
            StringBuilder sb = new StringBuilder();
            foreach (string value in values)
            {
                sb.Append(value);
                sb.Append('\n');
            }
            MessageBox.Show(sb.ToString());*/

            pdfFormFields.SetField("02", Rate);
            pdfFormFields.SetField("03", Surname);
            pdfFormFields.SetField("04", Name);
            pdfFormFields.SetField("05", SecondName);
            pdfFormFields.SetField("06", Birth.Day.ToString("D2"));
            pdfFormFields.SetField("07", Birth.Month.ToString("D2"));
            pdfFormFields.SetField("08", Birth.Year.ToString("D4"));
            pdfFormFields.SetField("09", Sex ? "1" : "0");
            pdfFormFields.SetField("10", Sex ? "0" : "1");
            pdfFormFields.SetField("13", Document);
            pdfFormFields.SetField("14", DocumentSerie);
            pdfFormFields.SetField("15", DocumentNumber);
            pdfFormFields.SetField("16", DocumentIssuedBy);
            pdfFormFields.SetField("18", DocumentIssueDate.Day.ToString("D2"));
            pdfFormFields.SetField("19", DocumentIssueDate.Month.ToString("D2"));
            pdfFormFields.SetField("20", DocumentIssueDate.Year.ToString("D4"));
            pdfFormFields.SetField("22", PlaceIndex);
            pdfFormFields.SetField("24", PlaceCity);
            pdfFormFields.SetField("25", PlaceStreet + ", " + PlaceBuilding + ", " + PlaceFlat);
            pdfFormFields.SetField("51", Number);
            pdfFormFields.SetField("58", SellerID);
            pdfFormFields.SetField("59", Seller);
            pdfFormFields.SetField("60", ICC);
            pdfFormFields.SetField("61", ICC_Suffix);
            pdfFormFields.SetField("62", Number);
            pdfFormFields.SetField("66", ContractConclusionDate.Day.ToString("D2"));
            pdfFormFields.SetField("67", ContractConclusionDate.Month.ToString("D2"));
            pdfFormFields.SetField("68", ContractConclusionDate.Year.ToString("D4"));
            pdfFormFields.SetField("69", ContractConclusionPlace);

            // flatten the form to remove editting options, set it to false
            // to leave the form open to subsequent manual edits

            pdfStamper.FormFlattening = false;
            // close the pdf

            pdfStamper.Close();
        }
    }
}
