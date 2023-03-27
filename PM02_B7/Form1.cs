using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;
using System.Reflection.Emit;
using System.Diagnostics;
using System.IO;

namespace PM02_B7
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            radioButton1.Checked = true;
            button1.Enabled = false;
        }
        public static bool triger_chek = false;
        public static string chek_text1;
        public static string chek_text2;
        public static string chek_text3;
        private void button2_Click(object sender, EventArgs e)  //Оформление квитанции
        {
            if (triger_chek == true)
            {
                // Создаём объект документа
                Microsoft.Office.Interop.Word.Document doc = null;
                try
                {
                    // Создаём объект приложения
                    Word.Application app = new Word.Application();
                    // Путь до шаблона документа
                    string source = System.IO.Path.GetFullPath("C:\\Users\\User\\Desktop\\PM02_B7\\чек.docx");
                    // Открываем
                    doc = app.Documents.Add(source);
                    doc.Activate();

                    // Добавляем информацию
                    // wBookmarks содержит все закладки
                    Microsoft.Office.Interop.Word.Bookmarks wBookmarks = doc.Bookmarks;
                    Word.Range wRange;
                    int i = 0;
                    Random random = new Random();
                    int randoms = random.Next(100000000, 999999999);
                    DateTime dateTime = DateTime.Now;
                    string[] data = new string[4] { $"{randoms}", $"{dateTime}", $"{chek_text3} (Товар:{chek_text1}{chek_text2} )", $"{rezult}" };
                    foreach (Word.Bookmark mark in wBookmarks)
                    {

                        wRange = mark.Range;
                        wRange.Text = data[i];
                        i++;
                    }

                    // Закрываем документ
                    doc.Close();
                    doc = null;

                    MessageBox.Show(
"Квитанция успешно сформирована!",
"Информация",
MessageBoxButtons.OK,
MessageBoxIcon.Information,
MessageBoxDefaultButton.Button1,
MessageBoxOptions.DefaultDesktopOnly);
                }
                catch (Exception ex)
                {
                    doc = null;
                    Console.WriteLine("Ошибка!");
                    Console.ReadLine();
                }
            }
            else
            {
                MessageBox.Show(
"Произошла ошибка! Квитанция не сформирована.",
"Ошибка",
MessageBoxButtons.OK,
MessageBoxIcon.Error,
MessageBoxDefaultButton.Button1,
MessageBoxOptions.DefaultDesktopOnly);
            }
        }
        public static double rezult;            //общая стоимость
        public static double cena = 0;          //цена за 1 билет
        int n;                                  //кол-во билетов
        
        private void button1_Click(object sender, EventArgs e)  //Расчет
        {

            if (radioButton1.Checked)
            {
                cena = 4000;
            }
            if (radioButton2.Checked)
            {
                cena = 2500;
            }
            if (radioButton3.Checked)
            {
                cena = 3500;
            }

            n = Convert.ToInt32(textBox1.Text);
            rezult = n * cena;

            label3.Text = "Цена билета: "+ cena.ToString("c") + "\nКоличество: " + n.ToString() + "шт. \n" + "Итого к оплате: " + rezult.ToString("c");
        }
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (textBox1.Text.Length == 0)
            {
                button1.Enabled = false;
            }
            else
                button1.Enabled = true;
            label3.Text = "";
        }
        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar >= '0') && (e.KeyChar <= '9'))
               if (Char.IsControl(e.KeyChar))
                {
                    if (e.KeyChar == (char)Keys.Enter)
                    {
                        button1.Focus();
                    }
                    return;
                }
            e.Handled = true;
        }
        public static string savePaths;
        public static string pathname;
        public static string opFILEsot1;
        private void button3_Click(object sender, EventArgs e)  //Добавление афиши
        {
            savePaths = System.IO.Path.GetFullPath("..\\Foto\\");
            OpenFileDialog OPF = new OpenFileDialog();
            OPF.Filter = "Изображения|*.png|*.jpeg|*.jpg";
            if (OPF.ShowDialog() == DialogResult.OK)
            {
                pathname = Path.GetFileName(OPF.FileName);
                opFILEsot1 = OPF.FileName;
                try
                {
                    pictureBox1.Image = Image.FromFile(OPF.FileName);
                    MessageBox.Show(
"Изображение загружено",
"Информация",
MessageBoxButtons.OK,
MessageBoxIcon.Information,
MessageBoxDefaultButton.Button1,
MessageBoxOptions.DefaultDesktopOnly);

                }
                catch
                {
                    MessageBox.Show(
"Не удалось загрузить изображене",
"Ошибка",
MessageBoxButtons.OK,
MessageBoxIcon.Error,
MessageBoxDefaultButton.Button1,
MessageBoxOptions.DefaultDesktopOnly);
                }

            }
            else
            {
                MessageBox.Show(
"Изображение не выбрано!",
"Внимание",
MessageBoxButtons.OK,
MessageBoxIcon.Warning,
MessageBoxDefaultButton.Button1,
MessageBoxOptions.DefaultDesktopOnly);
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            label3.Text = "";
            textBox1.Focus();
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            label3.Text = "";
            textBox1.Focus();
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            label3.Text = "";
            textBox1.Focus();
        }
    }
}
