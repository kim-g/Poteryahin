using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MTS_PDF_Table
{
    class PersonInfo
    {
        public string Number { get; set; }
        public string ICC { get; set; }
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
                        ICC = InTable.Rows[Position].ItemArray[Col.Ordinal].ToString();
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

            pdfFormFields.SetField("02", Rate);
            pdfFormFields.SetField("03", Surname);
            pdfFormFields.SetField("04", Name);
            pdfFormFields.SetField("05", SecondName);
            pdfFormFields.SetField("06", Birth.Day.ToString());
            pdfFormFields.SetField("07", Birth.Month.ToString());
            pdfFormFields.SetField("08", Birth.Year.ToString());
            pdfFormFields.SetField("09", Sex ? "Yes" : "0");
            pdfFormFields.SetField("10", Sex ? "0" : "Yes");
            pdfFormFields.SetField("13", Document);
            pdfFormFields.SetField("14", DocumentSerie);
            pdfFormFields.SetField("15", DocumentNumber);
            pdfFormFields.SetField("16", DocumentIssuedBy);
            pdfFormFields.SetField("18", DocumentIssueDate.Day.ToString());
            pdfFormFields.SetField("19", DocumentIssueDate.Month.ToString());
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

            // flatten the form to remove editting options, set it to false
            // to leave the form open to subsequent manual edits

            pdfStamper.FormFlattening = false;
            // close the pdf

            pdfStamper.Close();
        }
    }
}
