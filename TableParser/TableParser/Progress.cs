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

        public static PrBar Current = new PrBar();
        public static PrBar All = new PrBar();

        public static bool Abort = false;
    }

    public class PrBar
    {
        public int Maximum = 100;
        public int Position = 0;
        public int Done = 0;
    }

}
