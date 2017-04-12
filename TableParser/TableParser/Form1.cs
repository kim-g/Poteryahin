using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TableParser
{
    public partial class Form1 : Form
    {
        Config config = (Config)Serializer.LoadFromXML("config.xml", typeof(Config));

        public Form1()
        {
            InitializeComponent();
        }

        private void Licence()
        {
            MessageBox.Show("Программа написана Григорием Кимом (mail@kim-g.ru).\n\n" +
                "Программа распространяется по принципам «как есть» и «не стреляйте в пианиста, он играет, как умеет» по лицензии BSD\n\n" +
                @"Copyright(c) 2017, Grigory Kim
All rights reserved.

Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:

*Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.

* Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and / or other materials provided with the distribution.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS ''AS IS'' AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE."
                , "О программе");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            button1.Enabled = false;
            button4.Enabled = false;
            // Получаем данные о файлах
            string FileToOpen = Files_Directories.OpenFile("Открыть файл с исходными данными", "Файлы Excel (*.xlsx, *.xls)|*.xlsx;*.xls|Все файлы (*.*)|*.*");
            if (FileToOpen == null) return;
            string[] Filters = Files_Directories.OpenFiles("Открыть файлы с масками данных", "Файлы Excel (*.xlsx, *.xls)|*.xlsx;*.xls|Все файлы (*.*)|*.*");
            if (Filters == null) return;
            string OutDir = Files_Directories.OpenDirectory();
            if (OutDir == null) return;

            // Настройка ProgressBar
            Progress.Abort = false;
            Progress.All.Position = 0;
            Progress.All.Maximum = 1 + Filters.Count() + Filters.Count();
            Progress.Counting = true;

            // Загружаем файлы в RAM
            Excel_Table Data = Excel_Table.LoadFromFile(FileToOpen);
            if (Progress.Abort) return;
            Progress.All.Position++;
            List<Excel_Table> FilterTables = new List<Excel_Table>();
            for (int i = 0; i < Filters.Count(); i++)
            {
                FilterTables.Add(Excel_Table.LoadFromFile(Filters[i]));
                if (Progress.Abort) return;
                Progress.All.Position++;
            }

            

            // Обработка фильтров
            for (int i = 0; i < Filters.Count(); i++)
            {
                // Формируем список фильтров
                string FilterList = "";
                for (int j = 0; j < FilterTables[i].Table_Height; j++)
                    FilterList += FilterTables[i].list[0, j] + ";";
                if (Progress.Abort) return;
                Excel_Table Res = Data.CopyRows(FilterList, config.Colomn, config.HeadRows);
                if (Progress.Abort) return;
                Res.SaveToFile(OutDir+@"\"+ Path.GetFileNameWithoutExtension(Filters[i]) + "_OUT.xlsx");
                Progress.All.Position++;
            }

            Progress.Counting = false;
            MessageBox.Show("Задание выполнено");
            button1.Enabled = true;
            button4.Enabled = true;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (!Progress.Counting)
            {
                if (panel1.Visible) panel1.Visible = false;
                return;
            }

            if (PBCL.Text != Progress.Process) PBCL.Text = Progress.Process;

            double CurPosDouble = (double)Progress.Current.Position / (double)Progress.Current.Maximum * 1000f;
            int CurPos = (int)Math.Round(CurPosDouble);

            if (PBC.Value != CurPos) PBC.Value = CurPos;

            CurPosDouble = (double)Progress.All.Position / (double)Progress.All.Maximum * 1000f;
            CurPos = (int)Math.Round(CurPosDouble);

            if (PBA.Value != CurPos) PBA.Value = CurPos;

            if (!panel1.Visible) panel1.Visible = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Вы уверены, что хотите прервать процесс? Отменить это действие будет невозможно!", "Прервать процесс", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                Progress.Counting = false;
                Progress.Abort = true;
                button1.Enabled = true;
                button4.Enabled = true;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Licence();
        }
    }
}
