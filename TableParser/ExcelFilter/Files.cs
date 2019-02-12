using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

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
            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog()
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
        /// Открыть диалоговое окно открытия нескольких файлов
        /// </summary>
        /// <param name="Title">Заголовок окна</param>
        /// <param name="Filter">Фильтр файлов</param>
        /// <returns></returns>
        public static string[] OpenFiles(string Title, string Filter = "Файлы Excel 2007+ (*.xlsx)|*.xlsx|Все файлы (*.*)|*.*")
        {
            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog()
            {
                Filter = Filter,
                Title = Title,
                Multiselect = true,
                CheckPathExists = true,
                CheckFileExists = true
            };
            if (openFileDialog.ShowDialog() == true)
            {
                return openFileDialog.FileNames;
            }
            return null;
        }

        /// <summary>
        /// Открыть диалоговое окно открытия нескольких файлов
        /// </summary>
        /// <param name="Title">Заголовок окна</param>
        /// <param name="Filter">Фильтр файлов</param>
        /// <returns></returns>
        public static string OpenDirectory(string Title, string Filter = "Файлы Excel 2007+ (*.xlsx)|*.xlsx|Все файлы (*.*)|*.*")
        {
            FolderBrowserDialog folderBrowser = new FolderBrowserDialog();

            DialogResult result = folderBrowser.ShowDialog();

            if (string.IsNullOrWhiteSpace(folderBrowser.SelectedPath))
                return null;
            else
            {
                return folderBrowser.SelectedPath;
            }
        }

        /// <summary>
        /// Открыть диалоговое окно сохранения файла
        /// </summary>
        /// <param name="Title">Заголовок окна</param>
        /// <param name="Filter">Фильтр файлов</param>
        /// <returns></returns>
        public static string SaveFile(string Title, string Filter = "Файлы Excel 2007+ (*.xlsx)|*.xlsx|Все файлы (*.*)|*.*")
        {
            Microsoft.Win32.SaveFileDialog saveFileDialog = new Microsoft.Win32.SaveFileDialog()
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
