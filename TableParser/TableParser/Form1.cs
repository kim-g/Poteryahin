using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TableParser
{
    public partial class Form1 : Form
    {
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
            // Получаем данные о файлах
            string FileToOpen = Files_Directories.OpenFile("Открыть файл с исходными данными", "Файлы Excel (*.xlsx, *.xls)|*.xlsx;*.xls|Все файлы (*.*)|*.*");
            if (FileToOpen == null) return;
            string[] Filters = Files_Directories.OpenFiles("Открыть файлы с масками данных", "Файлы Excel (*.xlsx, *.xls)|*.xlsx;*.xls|Все файлы (*.*)|*.*");
            if (Filters == null) return;
            string OutDir = Files_Directories.OpenDirectory();
            if (OutDir == null) return;

            // Загружаем файлы в RAM
            Excel_Table Data = Excel_Table.LoadFromFile(FileToOpen);
            List<Excel_Table> FilterTables = new List<Excel_Table>();
            for (int i = 0; i < Filters.Count(); i++)
                FilterTables.Add(Excel_Table.LoadFromFile(Filters[i]));

            // Обработка фильтров
            for (int i = 0; i < Filters.Count(); i++)
            {
                // Формируем список фильтров
                string FilterList = "";
                for (int j = 0; j < FilterTables[i].Table_Height; j++)
                    FilterList += FilterTables[i].list[0, j] + ";";
                Excel_Table Res = Data.CopyRows(FilterList, 2);
                Res.SaveToFile(@"D:\Temp\Test.xlsx");
            }

            MessageBox.Show("Задание выполнено");
        }
    }
}
