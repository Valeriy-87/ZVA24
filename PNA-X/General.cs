using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace PNA_X
{
    public partial class General : Form
    {
        public static string predrp;
        public static string addres;
        public static string number_prot;
        public static string metodik_pover;

        public static string etalon_name1;
        public static string etalon_data1;


        public static string etalon_name2;
        public static string etalon_data2;


        public static string etalon_name3;
        public static string etalon_data3;


        public static string etalon_name4;
        public static string etalon_data4;


        public static string etalon_name5;
        public static string etalon_data5;

        public static string number_PO;
        public static string number_DK4;

        public bool isright = true;


        public General()
        {
            InitializeComponent();
        }

        private void button14_Click(object sender, EventArgs e)
        {

            isright = true;

            StreamWriter sw = new StreamWriter("General.txt");

            sw.WriteLine(textBox1.Text);
            sw.WriteLine(textBox3.Text);
            sw.WriteLine(textBox8.Text);

            sw.WriteLine(textBox5.Text);
            sw.WriteLine(dateTimePicker1.Text);
            if (dateTimePicker1.Value.AddYears(Convert.ToInt32(textBox31.Text)) < Form1.date || dateTimePicker1.Value > Form1.date)
            {
                textBox5.BackColor = Color.Red;
                isright = false;
            }
            else
            {
                textBox5.BackColor = Color.White;
            }

            sw.WriteLine(textBox6.Text);
            sw.WriteLine(dateTimePicker2.Text);
            if (dateTimePicker2.Value.AddYears(Convert.ToInt32(textBox30.Text)) < Form1.date || dateTimePicker2.Value > Form1.date)
            {
                textBox6.BackColor = Color.Red;
                isright = false;
            }
            else
            {
                textBox6.BackColor = Color.White;
            }

            sw.WriteLine(textBox9.Text);
            sw.WriteLine(dateTimePicker3.Text);
            if (dateTimePicker3.Value.AddYears(Convert.ToInt32(textBox29.Text)) < Form1.date || dateTimePicker3.Value > Form1.date)
            {
                textBox9.BackColor = Color.Red;
                isright = false;
            }
            else
            {
                textBox9.BackColor = Color.White;
            }

            sw.WriteLine(textBox11.Text);
            sw.WriteLine(dateTimePicker4.Text);
            if (dateTimePicker4.Value.AddYears(Convert.ToInt32(textBox28.Text)) < Form1.date || dateTimePicker4.Value > Form1.date)
            {
                textBox11.BackColor = Color.Red;
                isright = false;
            }
            else
            {
                textBox11.BackColor = Color.White;
            }

            sw.WriteLine(textBox17.Text);
            sw.WriteLine(dateTimePicker7.Text);
            if (dateTimePicker7.Value.AddYears(Convert.ToInt32(textBox25.Text)) < Form1.date || dateTimePicker7.Value > Form1.date)
            {
                textBox17.BackColor = Color.Red;
                isright = false;
            }
            else
            {
                textBox17.BackColor = Color.White;
            }

            sw.WriteLine(textBox20.Text);



            sw.WriteLine(textBox31.Text);
            sw.WriteLine(textBox30.Text);
            sw.WriteLine(textBox29.Text);
            sw.WriteLine(textBox28.Text);
            sw.WriteLine(textBox25.Text);

            sw.Close();

            FixResult();

            this.Close();
        }

        private void FixResult()
        {
            predrp = textBox1.Text;
            addres = textBox3.Text;
            number_prot = textBox8.Text;
            metodik_pover = textBox20.Text;

            etalon_name1 = textBox5.Text;
            etalon_data1 = dateTimePicker1.Value.AddYears(Convert.ToInt32(textBox31.Text)).AddDays(-1).ToShortDateString();

            etalon_name2 = textBox6.Text;
            etalon_data2 = dateTimePicker2.Value.AddYears(Convert.ToInt32(textBox30.Text)).AddDays(-1).ToShortDateString(); ;

            etalon_name3 = textBox9.Text;
            etalon_data3 = dateTimePicker3.Value.AddYears(Convert.ToInt32(textBox29.Text)).AddDays(-1).ToShortDateString(); ;

            etalon_name4 = textBox11.Text;
            etalon_data4 = dateTimePicker4.Value.AddYears(Convert.ToInt32(textBox28.Text)).AddDays(-1).ToShortDateString(); ;

            etalon_name5 = textBox17.Text;
            etalon_data5 = dateTimePicker7.Value.AddYears(Convert.ToInt32(textBox25.Text)).AddDays(-1).ToShortDateString();
        }

        private void General_Load(object sender, EventArgs e)
        {
            isright = true;

            StreamReader sr = new StreamReader("General.txt");

            textBox1.Text = sr.ReadLine();
            textBox3.Text = sr.ReadLine();
            textBox8.Text = sr.ReadLine();

            textBox5.Text = sr.ReadLine();
            dateTimePicker1.Value = Convert.ToDateTime(sr.ReadLine());


            textBox6.Text = sr.ReadLine();
            dateTimePicker2.Value = Convert.ToDateTime(sr.ReadLine());

            textBox9.Text = sr.ReadLine();
            dateTimePicker3.Value = Convert.ToDateTime(sr.ReadLine());

            textBox11.Text = sr.ReadLine();
            dateTimePicker4.Value = Convert.ToDateTime(sr.ReadLine());

            textBox17.Text = sr.ReadLine();
            dateTimePicker7.Value = Convert.ToDateTime(sr.ReadLine());

            textBox20.Text = sr.ReadLine();

            textBox31.Text = sr.ReadLine();
            textBox30.Text = sr.ReadLine();
            textBox29.Text = sr.ReadLine();
            textBox28.Text = sr.ReadLine();
            textBox25.Text = sr.ReadLine();


            if (dateTimePicker1.Value.AddYears(Convert.ToInt32(textBox31.Text)) < Form1.date || dateTimePicker1.Value > Form1.date)
            {
                isright = false;
                textBox5.BackColor = Color.Red;
            }
            if (dateTimePicker2.Value.AddYears(Convert.ToInt32(textBox30.Text)) < Form1.date || dateTimePicker2.Value > Form1.date)
            {
                isright = false;
                textBox6.BackColor = Color.Red;
            }
            if (dateTimePicker3.Value.AddYears(Convert.ToInt32(textBox29.Text)) < Form1.date || dateTimePicker3.Value > Form1.date)
            {
                isright = false;
                textBox9.BackColor = Color.Red;
            }
            if (dateTimePicker4.Value.AddYears(Convert.ToInt32(textBox28.Text)) < Form1.date || dateTimePicker4.Value > Form1.date)
            {
                isright = false;
                textBox11.BackColor = Color.Red;
            }
            if (dateTimePicker7.Value.AddYears(Convert.ToInt32(textBox25.Text)) < Form1.date || dateTimePicker7.Value > Form1.date)
            {
                isright = false;
                textBox17.BackColor = Color.Red;
            }

            sr.Close();

            FixResult();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void General_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (!isright)
            {
                DialogResult dr = MessageBox.Show("Обратите внимание, что средства поверки, выделеннные красным цветом, не поверены, продолжить?", "Внимание", MessageBoxButtons.YesNo);

                if (dr == DialogResult.Yes)         //если подтверждаем выход из программы
                {
                    Dispose();
                }
                else
                {
                    e.Cancel = true;
                }
            }
        }

        private void textBox31_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsDigit(e.KeyChar) && e.KeyChar != Convert.ToChar(8))
            {
                e.Handled = true;
            }
        }

        private void textBox23_MouseMove(object sender, MouseEventArgs e)
        {

        }

        private void textBox23_MouseLeave(object sender, EventArgs e)
        {
            if (((TextBox)sender).Text.Length == 0)
            {
                ((TextBox)sender).Text = "0";
            }

            else
            {
                int i = 0;

                while (((TextBox)sender).Text.Length != 1 && ((TextBox)sender).Text[0] == '0' && Char.IsDigit(((TextBox)sender).Text[1]))
                {
                    ((TextBox)sender).Text = ((TextBox)sender).Text.Remove(0, 1);
                    i++;
                }
            }
        }
    }
}
