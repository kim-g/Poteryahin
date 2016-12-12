/****************************************************************/
/*                                                              */
/*                     Модуль ввода строки                      */
/*                          Версия 1.0                          */
/*                                                              */
/*                     Автор – Григорий Ким                     */
/*                      kim-g@ios.uran.ru                       */
/*    Федеральное государственное бюджетное учреждение науки    */
/*     Институт органического синтеза им. И.Я. Постовского      */
/* Уральского отделения Российской академии наук (ИОС УрО РАН)  */
/*                                                              */
/*                 Распространяется на условиях                 */
/*            Berkeley Software Distribution license            */
/*                                                              */
/****************************************************************/

using System;
using System.Windows.Forms;

namespace Parser
{
    public partial class Input_String : Form    // Окно ввода текстовой строки
    {
        private string Res= "@Cancel@";            // Результат (@Cancel@ по умолчанию)

        public Input_String()
        {
            InitializeComponent();
        }

        public static string GetString(string Title, string Label, int Default=0) // Запрос текстовой строки извне
        {
            Input_String IS = new Input_String();

            IS.Text = Title;               // Поставить заголовок окна
            IS.label1.Text = Label;        // Поставить надпись перед полем ввода
            IS.numericUpDown1.Value = Default;    // Поставить значение по уморлчанию. 

            IS.ShowDialog();               // Показать модально

            return IS.Res;                 // Вернуть результат
        }

        private void button2_Click(object sender, EventArgs e)  // Если пользователь отменил
        {
            Res = "@Cancel@";           // Вернём "@Cancel@"
            Close();                    // И закроем окно
        }

        private void button1_Click(object sender, EventArgs e)  //Если пользователь нажал «OK»
        {
            Res = numericUpDown1.Value.ToString();        // Вернём то, что он ввёл
            Close();                    // И закроем окно
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
    }
}