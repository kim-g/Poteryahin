using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TableParser
{
    class Files_Directories
    {
        public static string OpenFile(string Title, string Filter= "Файлы Excel 2007+ (*.xlsx)|*.xlsx|Все файлы (*.*)|*.*")
        {
            using (var ofd = new OpenFileDialog())
            {
                ofd.Title = Title;
                ofd.Filter = Filter;
                ofd.AddExtension = true;
                ofd.Multiselect = false;
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    return ofd.FileName;
                }
                return null;
            }
        }

        public static string[] OpenFiles(string Title, string Filter = "Все файлы (*.*)|*.*")
        {
            using (var ofd = new OpenFileDialog())
            {
                ofd.Title = Title;
                ofd.Filter = Filter;
                ofd.AddExtension = true;
                ofd.Multiselect = true;
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    return ofd.FileNames;
                }
                return null;
            }
        }

        public static string OpenDirectory()
        {
            FolderBrowserDialog directoryDialog = new FolderBrowserDialog();
            directoryDialog.ShowDialog();
            return directoryDialog.SelectedPath == "" ? null : directoryDialog.SelectedPath;
        }
    }
}
