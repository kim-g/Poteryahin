using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace TableParser
{
    public class Excel_Table
    {

        public string[,] list;
        public int Table_Width;
        public int Table_Height;

        public Excel_Table(int Width, int Height)
        {
            Table_Width = Width;
            Table_Height = Height;
            list = new string[Table_Width, Table_Height];
        }

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
            
            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);// Находим последнюю ячейку.
            
            // Создаём новый Excel_Table объект
            Excel_Table ET = new Excel_Table(lastCell.Column, lastCell.Row);

            // Настройка прогрессбара
            Progress.Maximum = ET.Table_Width * ET.Table_Height + 2 * ET.Table_Height;
            Progress.Process = "Считывание данных из Excel";

            for (int i = 0; i < ET.Table_Width; i++) //по всем колонкам
                for (int j = 0; j < ET.Table_Height; j++) // по всем строкам
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

        private string[] RemoveDouble(string[] In)
        {
            var hashset = new HashSet<string>(); //Создаём объект типа множество.

            foreach (var x in In) // Проходимся по массиву и добавляем только те элементы, которых во множестве ещё нет
            {
                if (x == "") continue;
                if (!hashset.Contains(x))
                    hashset.Add(x);
            }

            Array.Resize(ref In, hashset.Count); // Изменяем размерность массива на необходимую
            In = hashset.ToArray(); // Перебрасываем элементы из множества обратно в массив
            return In;
        }

        public int RowsCount(string Filter="*", int Colomn=0)
        {
            if (Filter == "*") return Table_Height;

            int Count = 0;
            for (int i=0; i < Table_Height; i++)
                 if (list[Colomn, i] == Filter) Count++;
            return Count;
        }

        public Excel_Table CopyRows( string Filters = "*", int Colomn = 0)
        {
            string[] FilterList = Filters.Split(';');   // Разделяем фильтры
            FilterList = RemoveDouble(FilterList); // Удалим повторы

            List<int> RowsCopy = new List<int>();
            int Count = 0;
            foreach (string Filter in FilterList)
            {
                if (Filter == "*")
                {
                    Count += Table_Height;
                    continue;
                }

                for (int i = 0; i < Table_Height; i++)
                    for (int j = 0; j < Table_Width; j++)
                    {
                        if (list[Colomn, i] == Filter)
                        {
                            RowsCopy.Add(i);
                            break;
                        }
                    }
            }

            //Создаём новую таблицу
            Excel_Table FilteredTable = new Excel_Table(Table_Width, Count);

            int FT_Pos = 0;
            // и скопируем все подходящие данные в новую таблицу
            for (int i = 0; i < Table_Height; i++)
                for (int j = 0; j < FilterList.Count(); j++)
                    if (list[Colomn, i] == FilterList[j])
                    {
                        for (int k = 0; k < Table_Width; k++)
                            FilteredTable.list[k, FT_Pos] = list[k, i];
                        FT_Pos++;
                    }

            return FilteredTable;
        }


        public List<string> ListFromCell(int i, int j, char Spacer)
        {
            return (List<string>)list[i, j].Split(Spacer).ToList();
        }

        public void SaveToFile(string FileName)
        {
            //Открываем файл Экселя
            //Создаём приложение.
            Excel.Application ObjExcel = new Excel.Application();
            ObjExcel.SheetsInNewWorkbook = 1;
            //Создаём книгу.                                                                                                                                                        
            ObjExcel.Workbooks.Add(Type.Missing);
            //Получаем набор ссылок на объекты Workbook (на созданные книги)
            Excel.Workbooks excelappworkbooks;
            Excel.Workbook excelappworkbook;
            excelappworkbooks = ObjExcel.Workbooks;
            //Получаем ссылку на книгу 1 - нумерация от 1
            excelappworkbook = excelappworkbooks[1];
            //Запроса на сохранение для книги не должно быть
            excelappworkbook.Saved = true;
            // Формат 
            ObjExcel.DefaultSaveFormat = Excel.XlFileFormat.xlOpenXMLWorkbook;

            // Ищем нужные листы
            Excel.Sheets excelsheets;
            Excel.Worksheet excelworksheet;

            excelsheets = excelappworkbook.Worksheets;
            //Получаем ссылку на лист 1
            excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);
            Excel.Range excelcells = excelworksheet.get_Range("A1", Type.Missing);

            for (int i = 0; i < Table_Width; i++)
            {
                for (int j = 0; j < Table_Height; j++)
                {
                    excelcells.Value2 = list[i, j];
                    excelcells = excelcells.Offset[1, 0];
                }
                excelcells = excelcells.Offset[0 - Table_Width, 1];
            }
            excelappworkbook.SaveAs(FileName,  //object Filename
            Excel.XlFileFormat.xlOpenXMLWorkbook, //object FileFormat
            Type.Missing,                       //object Password 
            Type.Missing,                       //object WriteResPassword  
            Type.Missing,                       //object ReadOnlyRecommended
            Type.Missing,                       //object CreateBackup
            Excel.XlSaveAsAccessMode.xlNoChange,//XlSaveAsAccessMode AccessMode
            Type.Missing,                       //object ConflictResolution
            Type.Missing,                       //object AddToMru 
            Type.Missing,                       //object TextCodepage
            Type.Missing,                       //object TextVisualLayout
            Type.Missing);                      //object Local

            excelappworkbook.Close(false, Type.Missing, Type.Missing); //закрыть не сохраняя (уже сохранили)
            //Удаляем приложение (выходим из экселя) - а то будет висеть в процессах!
            ObjExcel.Quit();
        }
    }
}
