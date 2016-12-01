﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Parser
{
    public class Excel_Table
    {

        string[,] list;

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
            int Table_Width = lastCell.Column;
            int Table_Height = lastCell.Row;
            ET.list = new string[Table_Width, Table_Height]; // массив значений с листа равен по размеру листу
            for (int i = 0; i < Table_Width; i++) //по всем колонкам
                for (int j = 0; j < Table_Height; j++) // по всем строкам
                    ET.list[i, j] = ObjWorkSheet.Cells[j + 1, i + 1].Text.ToString();//считываем текст в строку
            ObjWorkBook.Close(false, Type.Missing, Type.Missing); //закрыть не сохраняя

            //Удаляем приложение (выходим из экселя) - а то будет висеть в процессах!
            ObjExcel.Quit();

            return ET;
        }
    }
}
