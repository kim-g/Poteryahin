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
        public static string OpenFile(string Filter= "Все файлы (*.*)|*.*")
        {
            using (var ofd = new OpenFileDialog())
            {
                ofd.Filter = Filter;
                ofd.AddExtension = true;
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    return ofd.FileName;
                }
                return "<Cancel>";
            }
        }

        public static string OpenDirectory()
        {
            FolderBrowserDialog directoryDialog = new FolderBrowserDialog();
            directoryDialog.ShowDialog();
            return directoryDialog.SelectedPath == "" ? "<Cancel>" : directoryDialog.SelectedPath;
        }
    }
}
