using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Parser
{
    public class Excel_Table
    {

        public string[,] list;
        public int Table_Width;
        public int Table_Height;

        public static Excel_Table LoadFromFile(string FileName)
        {
            //Открываем файл Экселя
            //Создаём приложение.
            Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            //Открываем книгу.                                                                                                                                                        
            Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(FileName, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            //Выбираем таблицу(лист).
            Excel.Worksheet ObjWorkSheet;
            ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];

            // Создаём новый Excel_Table объект
            Excel_Table ET = new Excel_Table();

            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);// Находим последнюю ячейку.
            ET.Table_Width = lastCell.Column;
            ET.Table_Height = lastCell.Row;

            // Настройка прогрессбара
            Progress.Maximum = ET.Table_Width * ET.Table_Height + 2 * ET.Table_Height;
            Progress.Process = "Считывание данных из Excel";

            ET.list = new string[ET.Table_Width, ET.Table_Height]; // массив значений с листа равен по размеру листу
            for (int i = 0; i < ET.Table_Width; i++) //по всем колонкам
                for (int j = 1; j < ET.Table_Height; j++) // по всем строкам
                {
                    ET.list[i, j] = ObjWorkSheet.Cells[j + 1, i + 1].Text.ToString();//считываем текст в строку
                    Progress.Position++;
                    Application.DoEvents();

                    if (Progress.Abort)
                    {
                        ObjWorkBook.Close(false, Type.Missing, Type.Missing); //закрыть не сохраняя

                        //Удаляем приложение (выходим из экселя) - а то будет висеть в процессах!
                        ObjExcel.Quit();
                        return ET;
                    }
                }
            ObjWorkBook.Close(false, Type.Missing, Type.Missing); //закрыть не сохраняя

            //Удаляем приложение (выходим из экселя) - а то будет висеть в процессах!
            ObjExcel.Quit();

            Progress.Done = ET.Table_Width * ET.Table_Height;

            return ET;
        }


        public List<string> ListFromCell(int i, int j, char Spacer)
        {
            return (List<string>)list[i, j].Split(Spacer).ToList();
        }
    }
}
