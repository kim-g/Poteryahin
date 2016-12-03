using System;
using System.Collections.Generic;
using System.IO;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Parser
{
    public partial class Form1 : Form
    {
        // Массив готовых XML
        List<CONTRACT> Contracts;

        public Form1()
        {
            InitializeComponent();
        }

        public static string OpenFile()
        {
            using (var ofd = new OpenFileDialog())
            {
                ofd.Filter = "Все файлы (*.*)|*.*";
                ofd.AddExtension = true;
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    return ofd.FileName;
                }
                return "<Cancel>";
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Создание массива
            Contracts = new List<CONTRACT>();

            // Загрузка таблицы
            string ExcelFileName = OpenFile();
            if (ExcelFileName == "<Cancel>") return;
            Excel_Table Table;
            try
            {
                Table = Excel_Table.LoadFromFile(ExcelFileName);
            }
            catch (Exception)
            {
                MessageBox.Show("Ошибка открытия файла «"+ ExcelFileName + "»", "Ошибка");
                return;
            }

            //Загрузка номеров строк
            Colomn_Numbers Colomn_N = (Colomn_Numbers)Serializer.LoadFromXML("Colomns.xml", typeof(Colomn_Numbers));

            // Загрузка констант из файла
            Editable_Params Params = (Editable_Params)Serializer.LoadFromXML("Parameters.xml", typeof(Editable_Params));

            // Вставка контента в объекты
            for (int i = 1; i < Table.Table_Height; i++)
                Contracts.Add(GetContract(Table, i, Params, Colomn_N));

            // Загрузка шаблона 
            string Example = File.ReadAllText("Example.xml", Encoding.GetEncoding("Windows-1251"));

            // Вставка данных в шаблон и сохранение этого безобразия.
            for (int i = 0; i < Contracts.Count; i++)
                File.WriteAllText("OUT\\Card_" + i.ToString("D8") + ".xml", Contracts[i].ToXMLString(Example), Encoding.GetEncoding("Windows-1251"));


        }

        private CONTRACT GetContract(Excel_Table Table, int i, Editable_Params Params, Colomn_Numbers Colomn_N)
        {
            CONTRACT Con = new CONTRACT();

            // Заполнение
            Con.Status = Params.Const.Status;
            Con.DealerCode = Params.Const.DealerCode;
            Con.DealerPointCode = Params.Const.DealerPointCode;
            Con.DealerContractCode = Params.Const.DealerContractCode;
            Con.DealerContractDate = Params.Const.DealerContractDate;
            Con.ABSContractCode = Params.Const.ABSContractCode;
            Con.BANKPROPLIST = Params.Const.BANKPROPLIST;
            Con.Comments = Params.Const.Comments;
            Con.CLIENTVER = Params.Const.CLIENTVER;

            //CUSTOMER
            Con.CUSTOMER.CUSTOMERTYPESId = Params.Const.CUSTOMERTYPESId;
            Con.CUSTOMER.SPHERESId = Params.Const.SPHERESId;
            Con.CUSTOMER.Resident = Params.Const.Resident;
            Con.CUSTOMER.Ratepayer = Params.Const.Ratepayer;

            //CUSTOMER.PERSON
            Con.CUSTOMER.PERSON.PERSONTYPESId = Params.Const.PERSONTYPESId;

            //CUSTOMER.PERSON.PERSONNAME
            Con.CUSTOMER.PERSON.PERSONNAME.SEXTYPESId = Table.list[Colomn_N.GENDER, i] == "м" ? "0" : "1"; // Проверить женский пол
            Con.CUSTOMER.PERSON.PERSONNAME.LastName = Table.list[Colomn_N.LAST_N, i];
            Con.CUSTOMER.PERSON.PERSONNAME.FirstName = Table.list[Colomn_N.FIRST_N, i];
            Con.CUSTOMER.PERSON.PERSONNAME.SecondName = Table.list[Colomn_N.PATRONYME, i];

            //CUSTOMER.PERSON.DOCUMENT
            Con.CUSTOMER.PERSON.DOCUMENT.DOCTYPESId = Params.GetID(Table.list[Colomn_N.DOCUMENTTYPE, i]);
            Con.CUSTOMER.PERSON.DOCUMENT.Seria = Table.list[Colomn_N.DOCUMENTID, i];
            Con.CUSTOMER.PERSON.DOCUMENT.Number = Table.list[Colomn_N.DOCUMENT_N, i];
            Con.CUSTOMER.PERSON.DOCUMENT.GivenBy = Table.list[Colomn_N.DOCISSUORIGINE, i];
            Con.CUSTOMER.PERSON.DOCUMENT.Date = Table.list[Colomn_N.DOCISSUDATE, i];
            Con.CUSTOMER.PERSON.DOCUMENT.Birthday = Table.list[Colomn_N.BIRTH_DATE, i];

            //CUSTOMER.PERSON
            Con.CUSTOMER.PERSON.INN = Params.Const.INN;

            //CUSTOMER.ADDRESS
            Con.CUSTOMER.ADDRESS.ZIP = Params.GetZIP(Table.list[Colomn_N.Region, i]);
            Con.CUSTOMER.ADDRESS.Country = Params.GetCountryName(Table.list[Colomn_N.COUNTRY, i]);
            Con.CUSTOMER.ADDRESS.Area = Params.GetRegion(Table.list[Colomn_N.Region, i]);
            Con.CUSTOMER.ADDRESS.Region = Params.Const.Region;
            Con.CUSTOMER.ADDRESS.PLACETYPESId = Params.GetCityID(Table.list[Colomn_N.PLACETYPE, i]);
            Con.CUSTOMER.ADDRESS.PlaceName = Table.list[Colomn_N.PLACENAMECITY, i];
            Con.CUSTOMER.ADDRESS.STREETTYPESId = Params.GetStreetID(Table.list[Colomn_N.STREETTYPE, i]);
            Con.CUSTOMER.ADDRESS.StreetName = Table.list[Colomn_N.STREETNAME, i];
            Con.CUSTOMER.ADDRESS.House = Table.list[Colomn_N.HOUSE_NO, i];
            Con.CUSTOMER.ADDRESS.BUILDINGTYPESId = Params.GetBuildingTypeID(Table.list[Colomn_N.BUILDINGTYPE, i]);
            Con.CUSTOMER.ADDRESS.Building = Params.GetBuildingID(Table.list[Colomn_N.BUILDING_NO, i]);
            Con.CUSTOMER.ADDRESS.ROOMTYPESId = Params.GetRoomTypeID(Table.list[Colomn_N.APARTTYPE, i]);
            Con.CUSTOMER.ADDRESS.Room = Table.list[Colomn_N.APPARTEMENT_NO, i];

            //DELIVERY
            Con.DELIVERY.DELIVERYTYPESId = Params.Const.DELIVERYTYPESId;
            Con.DELIVERY.Notes = Params.Const.Notes;

            //DELIVERY.ADDRESS
            Con.DELIVERY.ADDRESS.ZIP = Params.GetZIP(Table.list[Colomn_N.Region, i]);
            Con.DELIVERY.ADDRESS.Country = Params.GetCountryName(Table.list[Colomn_N.COUNTRY, i]);
            Con.DELIVERY.ADDRESS.Area = Params.GetRegion(Table.list[Colomn_N.Region, i]);
            Con.DELIVERY.ADDRESS.Region = Params.Const.Region;
            Con.DELIVERY.ADDRESS.PLACETYPESId = Params.GetCityID(Table.list[Colomn_N.PLACETYPE, i]);
            Con.DELIVERY.ADDRESS.PlaceName = Table.list[Colomn_N.PLACENAMECITY, i];
            Con.DELIVERY.ADDRESS.STREETTYPESId = Params.GetStreetID(Table.list[Colomn_N.STREETTYPE, i]);
            Con.DELIVERY.ADDRESS.StreetName = Table.list[Colomn_N.STREETNAME, i];
            Con.DELIVERY.ADDRESS.House = Table.list[Colomn_N.HOUSE_NO, i];
            Con.DELIVERY.ADDRESS.BUILDINGTYPESId = Params.GetBuildingTypeID(Table.list[Colomn_N.BUILDINGTYPE, i]);
            Con.DELIVERY.ADDRESS.Building = Params.GetBuildingID(Table.list[Colomn_N.BUILDING_NO, i]);
            Con.DELIVERY.ADDRESS.ROOMTYPESId = Params.GetRoomTypeID(Table.list[Colomn_N.APARTTYPE, i]);
            Con.DELIVERY.ADDRESS.Room = Table.list[Colomn_N.APPARTEMENT_NO, i];

            //CONTACT
            Con.CONTACT.PhonePrefix = Params.Const.PhonePrefix;
            Con.CONTACT.Phone = Params.Const.Phone;
            Con.CONTACT.FaxPrefix = Params.Const.FaxPrefix;
            Con.CONTACT.Fax = Params.Const.Fax;
            Con.CONTACT.EMail = Params.Const.EMail;
            Con.CONTACT.PagerOperatorPrefix = Params.Const.PagerOperatorPrefix;
            Con.CONTACT.PagerOperator = Params.Const.PagerOperator;
            Con.CONTACT.PagerAbonent = Params.Const.PagerAbonent;
            Con.CONTACT.Notes = Params.Const.Contact_Notes;

            //CONTACT.PERSONNAME
            Con.CONTACT.PERSONNAME.SEXTYPESId = Table.list[Colomn_N.GENDER, i] == "м" ? "0" : "1"; // Проверить женский пол
            Con.CONTACT.PERSONNAME.LastName = Table.list[Colomn_N.LAST_N, i] + " " + Table.list[Colomn_N.FIRST_N, i][0] + "." + Table.list[Colomn_N.PATRONYME, i][0] + ".";
            Con.CONTACT.PERSONNAME.FirstName = Params.Const.CP_FirstName;
            Con.CONTACT.PERSONNAME.SecondName = Params.Const.CP_SecondName;

            //CONNECTIONS.CONNECTION
            Con.CONNECTIONS.CONNECTION.PAYSYSTEMSId = Params.Const.PAYSYSTEMSId;
            Con.CONNECTIONS.CONNECTION.BILLCYCLESId = Params.Const.BILLCYCLESId;
            Con.CONNECTIONS.CONNECTION.CELLNETSId = Params.Const.CELLNETSId;
            Con.CONNECTIONS.CONNECTION.PRODUCTSId = Params.Const.PRODUCTSId;
            Con.CONNECTIONS.CONNECTION.PhoneOwner = Params.Const.PhoneOwner;
            Con.CONNECTIONS.CONNECTION.SerNumber = Params.Const.SerNumber;
            Con.CONNECTIONS.CONNECTION.SimLock = Params.Const.SimLock;
            Con.CONNECTIONS.CONNECTION.IMSI = Table.list[Colomn_N.IMSI, i];

            //CONNECTIONS.CONNECTION.MOBILES.MOBILE
            Con.CONNECTIONS.CONNECTION.MOBILES.MOBILE.CHANNELTYPESId = Params.Const.CHANNELTYPESId;
            Con.CONNECTIONS.CONNECTION.MOBILES.MOBILE.CHANNELLENSId = Params.Const.CHANNELLENSId;
            Con.CONNECTIONS.CONNECTION.MOBILES.MOBILE.SNB = Table.list[Colomn_N.CTN, i];
            Con.CONNECTIONS.CONNECTION.MOBILES.MOBILE.BILLPLANSId = Table.list[Colomn_N.BILLPLANSId, i];
            Con.CONNECTIONS.CONNECTION.MOBILES.MOBILE.SERVICES = Table.ListFromCell(Colomn_N.SERVICES, i, ' ');

            //LOGPARAMS
            Con.LOGPARAMS.AddRange(Params.Const.LOGPARAMS);

            return Con;
        }



        private void button2_Click(object sender, EventArgs e)
        {
            Editable_Params EP = (Editable_Params)Serializer.LoadFromXML("Parameters.xml", typeof(Editable_Params));

            // Сюда вставляем добавление параметров

            ///////////////////////////////////////

            EP.SaveToXML("Parameters.xml");
        }
    }
}
