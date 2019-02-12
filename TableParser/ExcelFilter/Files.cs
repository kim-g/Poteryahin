using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelFilter
{
    class Files
    {
        /// <summary>
        /// Открыть диалоговое окно открытия файла
        /// </summary>
        /// <param name="Title">Заголовок окна</param>
        /// <param name="Filter">Фильтр файлов</param>
        /// <returns></returns>
        public static string OpenFile(string Title, string Filter = "Файлы Excel 2007+ (*.xlsx)|*.xlsx|Все файлы (*.*)|*.*")
        {
            OpenFileDialog openFileDialog = new OpenFileDialog()
            {
                Filter=Filter,
                Title=Title,
                Multiselect=false,
                CheckPathExists=true,
                CheckFileExists = true
            };
            if (openFileDialog.ShowDialog() == true)
            {
                return openFileDialog.FileName;
            }
            return null;
        }

        /// <summary>
        /// Открыть диалоговое окно сохранения файла
        /// </summary>
        /// <param name="Title">Заголовок окна</param>
        /// <param name="Filter">Фильтр файлов</param>
        /// <returns></returns>
        public static string SaveFile(string Title, string Filter = "Файлы Excel 2007+ (*.xlsx)|*.xlsx|Все файлы (*.*)|*.*")
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog()
            {
                Filter = Filter,
                Title = Title,
                CheckPathExists = true,
                AddExtension = true
            };
            if (saveFileDialog.ShowDialog() == true)
            {
                return saveFileDialog.FileName;
            }
            return null;
        }
    }

}
