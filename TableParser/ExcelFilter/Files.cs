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

        public static string[] OpenFiles(string Title, string Filter = "Все файлы (*.*)|*.*")
        {
            return null;
        }

        public static string OpenDirectory()
        {
            return null;
        }
    }

}
