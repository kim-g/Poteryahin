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
        Colomn_Numbers Colomn_N = (Colomn_Numbers)Serializer.LoadFromXML("Colomns.xml", typeof(Colomn_Numbers));

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
            CONTRACT CN = new CONTRACT();
            CN.SaveToXML("Test.xml");

        }
    }
}
