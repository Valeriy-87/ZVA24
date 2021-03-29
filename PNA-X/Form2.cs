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
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            dataGridView1.RowTemplate.Height = 40;
            dataGridView1.Rows.Clear();

            dataGridView2.RowTemplate.Height = 40;
            dataGridView2.Rows.Clear();

            StreamReader sr = new StreamReader("att_et_abs_12.txt");
            dateTimePicker1.Value = Convert.ToDateTime(sr.ReadLine());

            while (!sr.EndOfStream)
            {
                string[] temp = sr.ReadLine().Split('\t');
                dataGridView1.Rows.Add(temp[0], temp[1], temp[2], temp[3]);
            }

            sr.Close();

            sr = new StreamReader("att_et_abs_21.txt");

            while (!sr.EndOfStream)
            {
                string[] temp = sr.ReadLine().Split('\t');
                dataGridView2.Rows.Add(temp[0], temp[1], temp[2], temp[3]);
            }

            sr.Close();

            if (dateTimePicker1.Value < Form1.date)
            {
                label1.ForeColor = Color.Red;
                MessageBox.Show("Обратите внимание, что на дату поверки просрочен срок обновления эталонных данных", "Внимание");
            }
            else
            {
                label1.ForeColor = Color.Black;
            }

            sr.Close();

        }

        private void button1_Click(object sender, EventArgs e)
        {
            StreamWriter sw = new StreamWriter("att_et_abs_12.txt");
            sw.WriteLine(dateTimePicker1.Value);

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                sw.WriteLine($"{dataGridView1.Rows[i].Cells[0].Value}\t{dataGridView1.Rows[i].Cells[1].Value}\t{dataGridView1.Rows[i].Cells[2].Value}\t{dataGridView1.Rows[i].Cells[3].Value}");
            }

            sw.Close();

            sw = new StreamWriter("att_et_abs_21.txt");

            for (int i = 0; i < dataGridView2.RowCount; i++)
            {
                sw.WriteLine($"{dataGridView2.Rows[i].Cells[0].Value}\t{dataGridView2.Rows[i].Cells[1].Value}\t{dataGridView2.Rows[i].Cells[2].Value}\t{dataGridView2.Rows[i].Cells[3].Value}");
            }

            sw.Close();

            this.Close();
        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            string temp = dataGridView2.Rows[e.RowIndex].Cells[e.ColumnIndex].FormattedValue.ToString();

            if (temp == "")
            {
                temp = "0";

                dataGridView2.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = temp;
            }

            else
            {
                int i = 0;

                while (temp.Length != 1 && temp[0] == '0' && Char.IsDigit(temp[1]))
                {
                    temp = temp.Remove(0, 1);
                    i++;
                }

                dataGridView2.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = temp;
            }


            if (temp[0] == ',' || temp[0] == '.')
            {
                temp = "0" + temp;
            }

            if (temp.IndexOf(".") != -1)
            {
                if (temp.Length == temp.IndexOf(".") + 1)
                {
                    temp += "0";
                }
            }

            if (temp.IndexOf(",") != -1)
            {
                if (temp.Length == temp.IndexOf(",") + 1)
                {
                    temp += "0";
                }
            }

            if (temp.IndexOf(".") != -1 || temp.IndexOf(",") != -1)
            {
                while (temp[temp.Length - 1] == '0' && (temp[temp.Length - 2] != '.' && temp[temp.Length - 2] != ','))
                {
                    temp = temp.Remove(temp.Length - 1, 1);
                }
            }

            dataGridView2.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = temp;
        }

        private void dataGridView1_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            e.Control.KeyPress -= new KeyPressEventHandler(Control_KeyPress);
            e.Control.KeyPress += new KeyPressEventHandler(Control_KeyPress);
        }

        void Control_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!(Char.IsDigit(e.KeyChar)) && !(((e.KeyChar == '.') || (e.KeyChar == ',')) && ((((TextBox)sender).Text.IndexOf(".") == -1) && (((TextBox)sender).Text.IndexOf(",") == -1))))
            {
                if (e.KeyChar != (char)Keys.Back)
                {
                    e.Handled = true;
                }
            }

            else if (e.KeyChar == ',')
            {
                e.KeyChar = '.';
            }
        }
    }
}

