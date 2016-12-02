using System;
using System.Collections.Generic;
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
            Constants Const = (Constants)Serializer.LoadFromXML("const.xml", typeof(Constants));

            // Вставка контента в объекты
            for (int i = 0; i < Table.Table_Height; i++)
                Contracts.Add(GetContract(Table, i, Const, Colomn_N));

            // Загрузка шаблона 
            string Example = System.IO.File.ReadAllText("Example.xml", Encoding.GetEncoding("Windows-1251"));

        }

        private CONTRACT GetContract(Excel_Table Table, int i, Constants Const, Colomn_Numbers Colomn_N)
        {
            CONTRACT Con = new CONTRACT();

            // Заполнение
            Con.Status = Const.Status;
            Con.DealerCode = Const.DealerCode;
            Con.DealerPointCode = Const.DealerPointCode;
            Con.DealerContractCode = Const.DealerContractCode;
            Con.DealerContractDate = Const.DealerContractDate;
            Con.ABSContractCode = Const.ABSContractCode;

            //CUSTOMER
            Con.CUSTOMER.CUSTOMERTYPESId = Const.CUSTOMERTYPESId;
            Con.CUSTOMER.SPHERESId = Const.SPHERESId;
            Con.CUSTOMER.Resident = Const.Resident;
            Con.CUSTOMER.Ratepayer = Const.Ratepayer;

            //CUSTOMER.PERSON
            Con.CUSTOMER.PERSON.PERSONTYPESId = Const.PERSONTYPESId;

            //CUSTOMER.PERSON.PERSONNAME
            Con.CUSTOMER.PERSON.PERSONNAME.SEXTYPESId = Table.list[Colomn_N.GENDER, i] == "м" ? "0" : "1"; // Проверить женский пол
            Con.CUSTOMER.PERSON.PERSONNAME.LastName = Table.list[Colomn_N.LAST_N, i];
            Con.CUSTOMER.PERSON.PERSONNAME.FirstName = Table.list[Colomn_N.FIRST_N, i];
            Con.CUSTOMER.PERSON.PERSONNAME.SecondName = Table.list[Colomn_N.PATRONYME, i];

            //CUSTOMER.PERSON.DOCUMENT
            //Con.CUSTOMER.PERSON.DOCUMENT

            return Con;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Constants Const = new Constants();
            Const.SaveToXML("const.xml");
        }
    }
}
