using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace MTS_PDF_Table
{
    /// <summary>
    /// Логика взаимодействия для Log.xaml
    /// </summary>
    public partial class Log : Window
    {
        public Log()
        {
            InitializeComponent();
            Visibility = Visibility.Hidden;
        }

        public void Add(string Text)
        {
            LogBox.Items.Add(Text);
            if (this.Visibility == Visibility.Hidden)
                this.Visibility = Visibility.Visible;
        }

        private void Copy_Click(object sender, RoutedEventArgs e)
        {
            string OutText = "";
            foreach (string X in LogBox.Items)
            {
                OutText += $"{X}\n";
            }
            Clipboard.SetText(OutText);
        }
    }
}
