using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TableParser
{
    // Хранит информацию о текущем процессе
    public class Progress
    {
        public static bool Counting = false;
        public static string Process = "";

        public static int Maximum = 100;
        public static int Position = 0;
        public static int Done = 0;

        public static bool Abort = false;
    }

}
