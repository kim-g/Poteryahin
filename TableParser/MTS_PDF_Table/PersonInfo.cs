using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Data;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

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
        public string BirthPlace { get; set; }
        public string Citizenship { get; set; }
        public string Document { get; set; }
        public string DocumentSerie { get; set; }
        public string DocumentNumber { get; set; }
        public string[] DocumentIssuedBy { get; set; } = new string[2];
        public DateTime DocumentIssueDate { get; set; }
        public string DocumentIndex { get; set; }
        public string[] DocumentIndex2 { get; set; }
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
        public string Line { get; set; }

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
            Line = (Position + MTS_PDF_Window.Head + 1).ToString("D6");
            foreach (DataColumn Col in InTable.Columns)
            {
                switch (Col.Caption.ToLowerInvariant())
                {
                    case "номер":
                        try
                        {
                            Number = InTable.Rows[Position].ItemArray[Col.Ordinal].ToString();
                            Number = Number.Replace("-", "").Replace("(", "").Replace(")", "");
                            Regex phone = new Regex(@"^((8|\+7)[\- ]?)?(\(?\d{3}\)?[\- ]?)?([\d\- ]{3})([\d\- ]{4,5})$");
                            Number = phone.Replace(Number, "$3$4$5");
                        }
                        catch (Exception e)
                        {
                            Number = "";
                            ToLog(InTable.TableName, Position, Col.Caption,e.Message);
                        }
                        break;
                    case "icc":
                        try
                        {
                            string ICC_Str = InTable.Rows[Position].ItemArray[Col.Ordinal].ToString().
                                Replace('-', ' ').Replace('_', ' ').Replace('=', ' ').Replace('+', ' ');
                            string[] ICC_Arr = ICC_Str.Split(' ');
                            ICC = ICC_Arr[0];
                            ICC_Suffix = ICC_Arr[1];
                        }
                        catch (Exception e)
                        {
                            ICC = "";
                            ICC_Suffix = "";
                            ToLog(InTable.TableName, Position, Col.Caption, e.Message);
                        }
                        break;
                    case "фамилия":
                        try
                        {
                            Surname = InTable.Rows[Position].ItemArray[Col.Ordinal].ToString();
                        }
                        catch (Exception e)
                        {
                            Surname = "";
                            ToLog(InTable.TableName, Position, Col.Caption, e.Message);
                        }
                        break;
                    case "имя":
                        try
                        {
                            Name = InTable.Rows[Position].ItemArray[Col.Ordinal].ToString();
                        }
                        catch (Exception e)
                        {
                            Name = "";
                            ToLog(InTable.TableName, Position, Col.Caption, e.Message);
                        }
                        break;
                    case "отчество":
                        try
                        {
                            SecondName = InTable.Rows[Position].ItemArray[Col.Ordinal].ToString();
                        }
                        catch (Exception e)
                        {
                            SecondName = "";
                            ToLog(InTable.TableName, Position, Col.Caption, e.Message);
                        }
                        break;
                    case "пол":
                        try
                        {
                            Sex = (InTable.Rows[Position].ItemArray[Col.Ordinal].ToString().ToLowerInvariant() == "м")
                            || (InTable.Rows[Position].ItemArray[Col.Ordinal].ToString().ToLowerInvariant() == "m");
                        }
                        catch (Exception e)
                        {
                            Sex = false;
                            ToLog(InTable.TableName, Position, Col.Caption, e.Message);
                        }
                        break;
                    case "место рождения":
                        try
                        {
                            BirthPlace = InTable.Rows[Position].ItemArray[Col.Ordinal].ToString();
                        }
                        catch (Exception e)
                        {
                            BirthPlace = "";
                            ToLog(InTable.TableName, Position, Col.Caption, e.Message);
                        }
                        break;
                    case "гражданство":
                        try
                        {
                            Citizenship = InTable.Rows[Position].ItemArray[Col.Ordinal].ToString();
                        }
                        catch (Exception e)
                        {
                            Citizenship = "";
                            ToLog(InTable.TableName, Position, Col.Caption, e.Message);
                        }
                        break;
                    case "документ удост личность":
                        try
                        {
                            Document = InTable.Rows[Position].ItemArray[Col.Ordinal].ToString();
                        }
                        catch (Exception e)
                        {
                            Document = "";
                            ToLog(InTable.TableName, Position, Col.Caption, e.Message);
                        }
                        break;
                    case "серия":
                        try
                        {
                            DocumentSerie = InTable.Rows[Position].ItemArray[Col.Ordinal].ToString();
                        }
                        catch (Exception e)
                        {
                            DocumentSerie = "";
                            ToLog(InTable.TableName, Position, Col.Caption, e.Message);
                        }
                        break;
                    case "номер документа":
                        try
                        {
                            DocumentNumber = InTable.Rows[Position].ItemArray[Col.Ordinal].ToString();
                        }
                        catch (Exception e)
                        {
                            DocumentNumber = "";
                            ToLog(InTable.TableName, Position, Col.Caption, e.Message);
                        }
                        break;
                    case "кем выдан":
                        try
                        {
                            string Str = InTable.Rows[Position].ItemArray[Col.Ordinal].ToString();
                            if (Str.Length < 50)
                            {
                                DocumentIssuedBy[0] = Str;
                                DocumentIssuedBy[1] = "";
                            }
                            else
                            {
                                for (int i=50; i>0; i--)
                                {
                                    if (Str[i] == ' ')
                                    {
                                        DocumentIssuedBy[0] = Str.Remove(i);
                                        DocumentIssuedBy[1] = Str.Substring(i + 1);
                                        break;
                                    }
                                }
                            }
                        }
                        catch (Exception e)
                        {
                            DocumentIssuedBy = new string[] { "", "" };
                            ToLog(InTable.TableName, Position, Col.Caption, e.Message);
                        }
                        break;
                    case "дата выдачи":
                        try
                        {
                            DocumentIssueDate = DateTime.Parse(InTable.Rows[Position].ItemArray[Col.Ordinal].ToString());
                        }
                        catch (Exception e)
                        {
                            DocumentIssueDate = DateTime.Now;
                            ToLog(InTable.TableName, Position, Col.Caption, e.Message);
                        }
                        break;
                    case "код подразделения":
                        try
                        {
                            DocumentIndex = InTable.Rows[Position].ItemArray[Col.Ordinal].ToString();
                            DocumentIndex2 = DocumentIndex.Split('-');
                        }
                        catch (Exception e)
                        {
                            DocumentIndex = "";
                            DocumentIndex2 = new string[] { "", "" };
                            ToLog(InTable.TableName, Position, Col.Caption, e.Message);
                        }
                        break;
                    case "дата рождения":
                        try
                        {
                            Birth = DateTime.Parse(InTable.Rows[Position].ItemArray[Col.Ordinal].ToString());
                        }
                        catch (Exception e)
                        {
                            Birth = DateTime.Now;
                            ToLog(InTable.TableName, Position, Col.Caption, e.Message);
                        }
                        break;
                    case "индекс":
                        try
                        {
                            PlaceIndex = InTable.Rows[Position].ItemArray[Col.Ordinal].ToString();
                        }
                        catch (Exception e)
                        {
                            PlaceIndex = "";
                            ToLog(InTable.TableName, Position, Col.Caption, e.Message);
                        }
                        break;
                    case "город":
                        try
                        {
                            PlaceCity = InTable.Rows[Position].ItemArray[Col.Ordinal].ToString();
                        }
                        catch (Exception e)
                        {
                            PlaceCity = "";
                            ToLog(InTable.TableName, Position, Col.Caption, e.Message);
                        }
                        break;
                    case "улица":
                        try
                        {
                            PlaceStreet = InTable.Rows[Position].ItemArray[Col.Ordinal].ToString();
                        }
                        catch (Exception e)
                        {
                            PlaceStreet = "";
                            ToLog(InTable.TableName, Position, Col.Caption, e.Message);
                        }
                        break;
                    case "дом":
                        try
                        {
                            PlaceBuilding = InTable.Rows[Position].ItemArray[Col.Ordinal].ToString();
                        }
                        catch (Exception e)
                        {
                            PlaceBuilding = "";
                            ToLog(InTable.TableName, Position, Col.Caption, e.Message);
                        }
                        break;
                    case "кв.":
                        try
                        {
                            PlaceFlat = InTable.Rows[Position].ItemArray[Col.Ordinal].ToString();
                        }
                        catch (Exception e)
                        {
                            PlaceFlat = "";
                            ToLog(InTable.TableName, Position, Col.Caption, e.Message);
                        }
                        break;
                    case "тариф":
                        try
                        {
                            Rate = InTable.Rows[Position].ItemArray[Col.Ordinal].ToString();
                        }
                        catch (Exception e)
                        {
                            Rate = "";
                            ToLog(InTable.TableName, Position, Col.Caption, e.Message);
                        }
                        break;
                    case "дата заключения догвора":
                        try
                        {
                            ContractConclusionDate = DateTime.Parse(InTable.Rows[Position].ItemArray[Col.Ordinal].ToString());
                        }
                        catch (Exception e)
                        {
                            ContractConclusionDate = DateTime.Now;
                            ToLog(InTable.TableName, Position, Col.Caption, e.Message);
                        }
                        break;
                    case "место заключения":
                        try
                        {
                            ContractConclusionPlace = InTable.Rows[Position].ItemArray[Col.Ordinal].ToString();
                        }
                        catch (Exception e)
                        {
                            ContractConclusionPlace = "";
                            ToLog(InTable.TableName, Position, Col.Caption, e.Message);
                        }
                        break;
                    case "продавец":
                        try
                        {
                            Seller = InTable.Rows[Position].ItemArray[Col.Ordinal].ToString();
                        }
                        catch (Exception e)
                        {
                            Seller = "";
                            ToLog(InTable.TableName, Position, Col.Caption, e.Message);
                        }
                        break;
                    case "код торговой точки":
                        try
                        {
                            SellerID = InTable.Rows[Position].ItemArray[Col.Ordinal].ToString();
                        }
                        catch (Exception e)
                        {
                            SellerID = "";
                            ToLog(InTable.TableName, Position, Col.Caption, e.Message);
                        }
                        break;
                }
            }
        }

        private void ToLog(string TableName, int Position, string ColumnName, string ErrorText)
        {
            MTS_PDF_Window.LogWindow.Add($"{TableName}: строка {Position + 2}, " +
                                $"столбец {ColumnName}: Ошибка чтения данных: {ErrorText}");
        }

        private string OnlyNum(string Text)
        {
            Regex ON = new Regex(@"/[^\d]+/g");
            return ON.Replace(Text, "");
        }

        private string charcode(int code)
        {
            string Res = CharCode(code);
            return Res.Substring(0, 1);
        }

        private string CharCode(int code)
        {
            ToLog("CharCode", 0, "", code.ToString());

            switch (code)
            {
                case 0: return new string(Encoding.Unicode.GetChars(BitConverter.GetBytes(8364)));
                case 95: return new string(Encoding.Unicode.GetChars(BitConverter.GetBytes(8216)));
                case 96: return new string(Encoding.Unicode.GetChars(BitConverter.GetBytes(8217)));
                case 97: return new string(Encoding.Unicode.GetChars(BitConverter.GetBytes(8220)));
                case 98: return new string(Encoding.Unicode.GetChars(BitConverter.GetBytes(8221)));
                case 99: return new string(Encoding.Unicode.GetChars(BitConverter.GetBytes(8226)));
                case 100: return new string(Encoding.Unicode.GetChars(BitConverter.GetBytes(8211)));
                case 101: return new string(Encoding.Unicode.GetChars(BitConverter.GetBytes(8212)));
                case 102: return new string(Encoding.Unicode.GetChars(BitConverter.GetBytes(732)));
                default: return new string(Encoding.Unicode.GetChars(BitConverter.GetBytes(code + 32)));
            }
        }

        public void FillForm(string newFile)
        {
            var pdfFile = newFile;

            using (var templateReader = new PdfReader(pdfTemplate))
            {
                using (var resultStamper = new PdfStamper(templateReader,
                    new FileStream(pdfFile, FileMode.Create)))
                {
                    // Получаем ссылку на форму с полями.
                    var form = resultStamper.AcroFields;
                    // Получаем все шрифты формы.
                    var fonts = templateReader.GetFormFonts();

                    // Установка значений полей формы.
                    form.SetFieldWithFont(templateReader, fonts, "02", Rate);
                    form.SetFieldWithFont(templateReader, fonts, "03", Surname);
                    form.SetFieldWithFont(templateReader, fonts, "04", Name);
                    form.SetFieldWithFont(templateReader, fonts, "05", SecondName);
                    form.SetFieldWithFont(templateReader, fonts, "06", Birth.Day.ToString("D2"));
                    form.SetFieldWithFont(templateReader, fonts, "07", Birth.Month.ToString("D2"));
                    form.SetFieldWithFont(templateReader, fonts, "08", Birth.Year.ToString("D4"));
                    form.SetField("09", Sex ? "1" : "0");
                    form.SetField("10", Sex ? "0" : "1");
                    form.SetFieldWithFont(templateReader, fonts, "11", BirthPlace);
                    form.SetFieldWithFont(templateReader, fonts, "13", Citizenship);
                    form.SetFieldWithFont(templateReader, fonts, "12", Document);
                    form.SetFieldWithFont(templateReader, fonts, "14", DocumentSerie);
                    form.SetFieldWithFont(templateReader, fonts, "15", DocumentNumber);
                    form.SetFieldWithFont(templateReader, fonts, "16", DocumentIssuedBy[0] + DocumentIssuedBy[1]);
                    //form.SetFieldWithFont(templateReader, fonts, "17", DocumentIssuedBy[1]);
                    form.SetFieldWithFont(templateReader, fonts, "17", DocumentIssueDate.Day.ToString("D2"));
                    form.SetFieldWithFont(templateReader, fonts, "18", DocumentIssueDate.Month.ToString("D2"));
                    form.SetFieldWithFont(templateReader, fonts, "19", DocumentIssueDate.Year.ToString("D4"));
                    form.SetFieldWithFont(templateReader, fonts, "20", DocumentIndex2[0]);
                    form.SetFieldWithFont(templateReader, fonts, "21", PlaceIndex);
                    form.SetFieldWithFont(templateReader, fonts, "22", DocumentIndex2[1]);
                    form.SetFieldWithFont(templateReader, fonts, "25", PlaceCity);
                    form.SetFieldWithFont(templateReader, fonts, "26", PlaceStreet + ", " + PlaceBuilding + ", " + PlaceFlat);
                    form.SetFieldWithFont(templateReader, fonts, "49", SellerID);
                    form.SetFieldWithFont(templateReader, fonts, "51", Seller);
                    form.SetFieldWithFont(templateReader, fonts, "iccid", ICC + "-" + ICC_Suffix);
                    form.GetRevision("iccid");
                    form.SetFieldWithFont(templateReader, fonts, "phone", Number);
                    form.GetRevision("phone");
                    form.SetFieldWithFont(templateReader, fonts, "60", ContractConclusionDate.Day.ToString("D2"));
                    form.SetFieldWithFont(templateReader, fonts, "61", ContractConclusionDate.Month.ToString("D2"));
                    form.SetFieldWithFont(templateReader, fonts, "62", ContractConclusionDate.Year.ToString("D4"));
                    form.SetFieldWithFont(templateReader, fonts, "63", ContractConclusionPlace);

                    string Text = OnlyNum(ICC + ICC_Suffix + Number);
                    int check = 105;
                    string result = "";
                    for (var i = 0; i < Text.Length / 2; i++)
                    {
                        var substr = Text.Substring(i * 2, 2);
                        int intval = Convert.ToInt32(substr);
                        check = check + intval * (i + 1);
                        result = result + charcode(intval);
                    }
                    check = check % 103;
                    result = "›" + result + charcode(check) + "œ";

                    ToLog("CharCode", 0, "", result);

                    form.SetField("barcode", result);

                    // Установка запрета на редактирование полей.
                    //resultStamper.FormFlattening = true;

                    resultStamper.Close();
                }
                templateReader.Close();
            }
            
            /*
            FontFactory.RegisterDirectory(Environment.GetFolderPath(Environment.SpecialFolder.Fonts));

            PdfReader pdfReader = new PdfReader(pdfTemplate);
            PdfStamper pdfStamper = new PdfStamper(pdfReader, new FileStream(
                newFile, FileMode.Create));
            AcroFields pdfFormFields = pdfStamper.AcroFields;

            pdfFormFields.AddSubstitutionFont(BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.WINANSI, false));

            //pdfFormFields.SetField("02",  Rate, Rate, true);
            //pdfFormFields.SetField("03", Surname);
            pdfFormFields.SetField("04", Name);
            pdfFormFields.SetField("05", SecondName);
            pdfFormFields.SetField("06", Birth.Day.ToString("D2"));
            pdfFormFields.SetField("07", Birth.Month.ToString("D2"));
            pdfFormFields.SetField("08", Birth.Year.ToString("D4"));
            pdfFormFields.SetField("09", Sex ? "1" : "0");
            pdfFormFields.SetField("10", Sex ? "0" : "1");
            pdfFormFields.SetField("11", BirthPlace);
            pdfFormFields.SetField("13", Citizenship);
            pdfFormFields.SetField("12", Document);
            pdfFormFields.SetField("14", DocumentSerie);
            pdfFormFields.SetField("15", DocumentNumber);
            pdfFormFields.SetField("16", DocumentIssuedBy[0] + DocumentIssuedBy[1]);
            //pdfFormFields.SetField("17", DocumentIssuedBy[1]);
            pdfFormFields.SetField("17", DocumentIssueDate.Day.ToString("D2"));
            pdfFormFields.SetField("18", DocumentIssueDate.Month.ToString("D2"));
            pdfFormFields.SetField("19", DocumentIssueDate.Year.ToString("D4"));
            pdfFormFields.SetField("20", DocumentIndex2[0]);
            pdfFormFields.SetField("21", PlaceIndex);
            pdfFormFields.SetField("22", DocumentIndex2[1]);
            pdfFormFields.SetField("25", PlaceCity);
            pdfFormFields.SetField("26", PlaceStreet + ", " + PlaceBuilding + ", " + PlaceFlat);
            pdfFormFields.SetField("49", SellerID);
            pdfFormFields.SetField("51", Seller);
            pdfFormFields.SetField("iccid", ICC+"-"+ ICC_Suffix);
            pdfFormFields.SetField("phone", Number);
            pdfFormFields.SetField("60", ContractConclusionDate.Day.ToString("D2"));
            pdfFormFields.SetField("61", ContractConclusionDate.Month.ToString("D2"));
            pdfFormFields.SetField("62", ContractConclusionDate.Year.ToString("D4"));
            pdfFormFields.SetField("63", ContractConclusionPlace);

            // flatten the form to remove editting options, set it to false
            // to leave the form open to subsequent manual edits

            pdfStamper.FormFlattening = false;
            // close the pdf

            pdfStamper.Close();*/
        }
    }
}
