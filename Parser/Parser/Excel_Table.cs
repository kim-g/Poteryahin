using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;

namespace Parser
{
    public class Excel_Table
    {

        public string[,] list;
        public int Table_Width;
        public int Table_Height;

        public static string[,] ThreadList;
        public static int Thread_Table_Height;
        public static Excel.Worksheet Thread_ObjWorkSheet;
        public static bool[] Stop;

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
            Thread_ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];

            // Создаём новый Excel_Table объект
            Excel_Table ET = new Excel_Table();

            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);// Находим последнюю ячейку.
            ET.Table_Width = lastCell.Column;
            ET.Table_Height = lastCell.Row;
            ET.list = new string[ET.Table_Width, ET.Table_Height]; // массив значений с листа равен по размеру листу
            ThreadList = new string[ET.Table_Width, ET.Table_Height]; // массив значений с листа равен по размеру листу
            Thread_Table_Height = ET.Table_Height;
            Stop = new bool[Thread_Table_Height];
            for (int i = 0; i < Stop.Count(); i++) Stop[i] = false;

            List<Thread> threads = new List<Thread>();
            for (int i = 0; i < ET.Table_Width; i++) //по всем колонкам
            {
                /*threads.Add(new Thread(CopyInThread));
                threads[i].Start(i);*/
                ThreadPool.QueueUserWorkItem(CopyInThread, i);
            }

            bool AllStop = false;
            do
            {
                Thread.Sleep(500);
                AllStop = true;
                for (int i = 0; i < ET.Table_Width; i++) //по всем колонкам
                    AllStop = AllStop && Stop[i]; // Проверим, что копирование завершилось.
            } while (!AllStop);

            ObjWorkBook.Close(false, Type.Missing, Type.Missing); //закрыть не сохраняя

            //Удаляем приложение (выходим из экселя) - а то будет висеть в процессах!
            ObjExcel.Quit();

            return ET;
        }

        private static void CopyInThread(object Params)
        {
            // Получаем функции извне
            int i = (int)Params;

            // Проделываем операцию с одним столбцом
            for (int j = 1; j < Thread_Table_Height; j++) // по всем строкам
                ThreadList[i, j] = Thread_ObjWorkSheet.Cells[j + 1, i + 1].Text.ToString();//считываем текст в строку

            // Сообщаем, что всё готово
            
        }


        public List<string> ListFromCell(int i, int j, char Spacer)
        {
            return (List<string>)list[i, j].Split(Spacer).ToList();
        }
    }
}
