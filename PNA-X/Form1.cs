using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Ivi.Driver.Interop;
using System.Threading;
using System.Diagnostics;
using System.IO;
using System.Globalization;
using PNA_X.Properties;
using System.Security.Cryptography;
using System.Management;
using NationalInstruments.VisaNS;

namespace PNA_X
{
    public partial class Form1 : Form, DrawGraphicOnForm
    {
        NumberFormatInfo provider = new NumberFormatInfo() { NumberDecimalSeparator = "." };

        public static DateTime date;
        Form3 fm2;
        Form2 fm3;
        Form4 fm4;

        public static StringBuilder buff;

        public static string addr_zva;
        public static string addr_CNT;

        public static string path;

        Calculate cl;

        private object m_objLock = new object();
        private bool m_drawing = false;

        General fm;

        private bool IsConnectN5242 = false;
        private bool IsConnectCNT = false;
        public static bool IsRun = false;

        private MessageBasedSession CNT;
        private MessageBasedSession VAC;

        public static int count_port_freq;
        public static double[] freqList_freq;
        public static int[] number_port_freq;

        public static int count_port_noise;
        public static int[] number_port_noise;


        public Form1()
        {
            InitializeComponent();
        }

        private delegate void UpdateEndCalculateDelegate();

        public void EndCalculate()
        {
            if (!this.InvokeRequired)
            {
                #region Разблокировка компонент формы

                button2.Enabled = true;
                button5.Enabled = true;
                button6.Enabled = true;
                button17.Enabled = true;

                groupBox5.Enabled = true;
                groupBox2.Enabled = true;
              


                progressBar1.Visible = false;

                button13.Enabled = false;

                #endregion
            }
            else
            {
                //Метод вызван из вторичного потока
                //Вот так вызываем метод, для того чтобы он выполнялся в главном потоке
                UpdateEndCalculateDelegate updDelegate = new UpdateEndCalculateDelegate(EndCalculate);
                Invoke(updDelegate);
            }
        }

        private void ShowGraphicsInternal()
        {
            lock (m_objLock)
            {
                m_drawing = false;
            }

            progressBar1.Value++;
        }

        public void DrawGraphic(bool p_endCalculate)
        {
            //Вызывается из вторичного потока
            ShowGraphics(p_endCalculate);
        }

        private delegate void UpdateNewValuesDelegate();

        private void ShowGraphics(bool p_endCalculate)
        {
            if (p_endCalculate == false)
            {
                //Это не конец расчета - проверяем идет ли уже рисование
                lock (m_objLock)
                {
                    if (m_drawing == true)
                    {
                        //Уже идет рисование - в очередь не добавляем промежуточный отчет
                        return;
                    }
                    else m_drawing = true;
                }
            }

            UpdateNewValuesDelegate updDelegate = new UpdateNewValuesDelegate(ShowGraphicsSynch);
            IAsyncResult result = updDelegate.BeginInvoke(null, null);
        }

        private void ShowGraphicsSynch()
        {
            if (!this.InvokeRequired)
            {
                //Метод вызван из главного потока
                ShowGraphicsInternal();
            }
            else
            {
                //Метод вызван из вторичного потока
                //Вот так вызываем метод, для того чтобы он выполнялся в главном потоке
                UpdateNewValuesDelegate updDelegate = new UpdateNewValuesDelegate(ShowGraphicsSynch);
                Invoke(updDelegate, new object[] { });
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (!button1.Enabled)
            {
                return;
            }

            if (!IsConnectCNT || !IsConnectN5242)
            {
                System.Media.SystemSounds.Beep.Play(); System.Media.SystemSounds.Beep.Play(); System.Media.SystemSounds.Beep.Play();
                MessageBox.Show("Проверьте физическое и программное подключение частотомера и анализатора цепей к компьютеру", "Внимание", MessageBoxButtons.OK);
                return;
            }

            DialogResult dr;

            count_port_freq = 0;

            number_port_freq = new int[4];

            int k = 0;

            if (checkBox1.Checked)
            {
                count_port_freq++;
                number_port_freq[k] = 1;
                k++;
            }

            if (checkBox2.Checked)
            {
                count_port_freq++;
                number_port_freq[k] = 2;
                k++;
            }

            if (checkBox3.Checked)
            {
                count_port_freq++;
                number_port_freq[k] = 3;
                k++;
            }

            if (checkBox4.Checked)
            {
                count_port_freq++;
                number_port_freq[k] = 4;
            }

            IsRun = false;

            try
            {

                path = CreateFolder(textBox6.Text);

                if (path == String.Empty)
                {
                    System.Media.SystemSounds.Beep.Play(); System.Media.SystemSounds.Beep.Play(); System.Media.SystemSounds.Beep.Play();
                    MessageBox.Show("Измените расположение файла поверки. Затем начните поверку заново", "Внимание");
                    return;
                }

                button2.Enabled = false;
                button5.Enabled = false;
                button6.Enabled = false;
                button17.Enabled = false;

                groupBox5.Enabled = false;
                groupBox2.Enabled = false;
           


                button13.Enabled = true;

                progressBar1.Visible = false;

                freqList_freq = new double[] { 0.01, 0.1, 1, 4, 5, 7, 10, 13, 15, 18, 20, 22, 24 };

                lock (m_objLock)
                {
                    m_drawing = false;
                }

                progressBar1.Minimum = 0;
                progressBar1.Maximum = freqList_freq.Length * count_port_freq - 1;
                progressBar1.Value = 0;

                progressBar1.Visible = true;

                cl.Start(this);
            }

            catch
            {
                System.Media.SystemSounds.Beep.Play(); System.Media.SystemSounds.Beep.Play(); System.Media.SystemSounds.Beep.Play();
                MessageBox.Show("Нарушение инструкции измерений, повторите попытку поверки по частоте еще раз", "Внимание");
                EndCalculate();
            }
        }


        private string CreateFolder(string str)
        {
            try
            {
                if (!Directory.Exists(str))
                {
                    Directory.CreateDirectory(str);
                }

                string[] str1 = VAC.Query("*IDN?").Split(',');


                if (!Directory.Exists(str + "\\" + str1[1]))
                {
                    Directory.CreateDirectory(str + "\\" + str1[1]);
                }

                if (!Directory.Exists(str + "\\" + str1[1] + "\\" + str1[2].Substring(str1[2].Length - 6, 6)))
                {
                    Directory.CreateDirectory(str + "\\" + str1[1] + "\\" + str1[2].Substring(str1[2].Length - 6, 6));
                }

                if (!Directory.Exists(str + "\\" + str1[1] + "\\" + str1[2].Substring(str1[2].Length - 6, 6) + "\\" + DateTime.Now.Date.ToShortDateString()))
                {
                    Directory.CreateDirectory(str + "\\" + str1[1] + "\\" + str1[2].Substring(str1[2].Length - 6, 6) + "\\" + DateTime.Now.Date.ToShortDateString());
                }

                return str + "\\" + str1[1] + "\\" + str1[2].Substring(str1[2].Length - 6, 6) + "\\" + DateTime.Now.Date.ToShortDateString();
            }

            catch
            {
                return String.Empty;
            }
        }

        private void button5_Click_1(object sender, EventArgs e)
        {
            if (!button5.Enabled)
            {
                return;
            }

            textBox1.BackColor = Color.Red;
            pictureBox6.BackgroundImage = Resources.rojo__verde_iconos_ok_ok_no_17_1106090017;
            IsConnectN5242 = false;

            try
            {

                addr_zva = textBox1.Text.Trim();

                // Create driver instance
                VAC = (MessageBasedSession)ResourceManager.GetLocalManager().Open(addr_zva);

                VAC.Write("*RST");
                VAC.Write("SYST:DISP:UPD ON");

                pictureBox6.BackgroundImage = Resources.OK_symbol;
                textBox1.BackColor = Color.White;
                IsConnectN5242 = true;

                StreamWriter sw = new StreamWriter("adress_zva.txt");
                sw.WriteLine(textBox1.Text);
                sw.Close();
            }
            catch (Exception ex)
            {
                pictureBox6.BackgroundImage = Resources.rojo__verde_iconos_ok_ok_no_17_1106090017;
                IsConnectN5242 = false;
                textBox1.BackColor = Color.Red;
                textBox1.Focus();
            }
        }


        private void button2_Click(object sender, EventArgs e)
        {
            if (!button2.Enabled)
            {
                return;
            }

            textBox2.BackColor = Color.Red;
            pictureBox2.BackgroundImage = Resources.rojo__verde_iconos_ok_ok_no_17_1106090017;
            IsConnectCNT = false;

            try
            {
                addr_CNT = textBox2.Text.Trim();

                // Create driver instance
                CNT = (MessageBasedSession)ResourceManager.GetLocalManager().Open(addr_CNT);
                pictureBox2.BackgroundImage = Resources.OK_symbol;
                textBox2.BackColor = Color.White;
                IsConnectCNT = true;

                StreamWriter sw = new StreamWriter("adress_cnt.txt");
                sw.WriteLine(textBox2.Text);
                sw.Close();
            }
            catch
            {
                pictureBox2.BackgroundImage = Resources.rojo__verde_iconos_ok_ok_no_17_1106090017;
                IsConnectCNT = false;
                textBox2.BackColor = Color.Red;
                textBox2.Focus();
            }
        }

        private void textBox6_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = true;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (!button6.Enabled)
            {
                return;
            }

            DialogResult dr = folderBrowserDialog1.ShowDialog();

            if (dr == DialogResult.OK)
            {
                textBox6.Text = folderBrowserDialog1.SelectedPath;

                StreamWriter sw = new StreamWriter("path.txt");

                sw.WriteLine(textBox6.Text);

                sw.Close();
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            if (!button9.Enabled)
            {
                return;
            }

            if (!IsConnectN5242)
            {
                System.Media.SystemSounds.Beep.Play(); System.Media.SystemSounds.Beep.Play(); System.Media.SystemSounds.Beep.Play();
                MessageBox.Show("Проверьте физическое и программное подключение анализатора цепей к компьютеру", "Внимание", MessageBoxButtons.OK);
                return;
            }

            count_port_noise = 0;

            number_port_noise = new int[4];

            int k = 0;

            if (checkBox15.Checked)
            {
                count_port_noise++;
                number_port_noise[k] = 1;
                k++;
            }

            if (checkBox16.Checked)
            {
                count_port_noise++;
                number_port_noise[k] = 2;
                k++;
            }

            if (checkBox17.Checked)
            {
                count_port_noise++;
                number_port_noise[k] = 3;
                k++;
            }

            if (checkBox14.Checked)
            {
                count_port_noise++;
                number_port_noise[k] = 4;
            }

            IsRun = false;

            try
            {
                path = CreateFolder(textBox6.Text);

                if (path == String.Empty)
                {
                    System.Media.SystemSounds.Beep.Play(); System.Media.SystemSounds.Beep.Play(); System.Media.SystemSounds.Beep.Play();
                    MessageBox.Show("Измените расположение файла поверки. Затем начните поверку заново", "Внимание");
                    return;
                }

                button2.Enabled = false;
                button5.Enabled = false;
                button6.Enabled = false;
                button17.Enabled = false;

                groupBox5.Enabled = false;
                groupBox2.Enabled = false;
             

                button13.Enabled = true;

                progressBar1.Visible = false;

                lock (m_objLock)
                {
                    m_drawing = false;
                }

                progressBar1.Minimum = 0;
                progressBar1.Maximum = count_port_noise;
                progressBar1.Value = 0;

                progressBar1.Visible = true;

                cl.Start1(this);
            }

            catch (Exception ex)
            {
                System.Media.SystemSounds.Beep.Play(); System.Media.SystemSounds.Beep.Play(); System.Media.SystemSounds.Beep.Play();
                MessageBox.Show("Нарушение инструкции измерений, повторите попытку поверки по уровню шумов еще раз", "Внимание");
            }
        }

        private void Form1_MouseDown(object sender, MouseEventArgs e)
        {
            this.Capture = false;
            Message n = Message.Create(this.Handle, 0xa1, new IntPtr(2), IntPtr.Zero);
            this.WndProc(ref n);
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            //появление диалогового окна при выходе из прграммы
            DialogResult dr;
            System.Media.SystemSounds.Beep.Play(); System.Media.SystemSounds.Beep.Play(); System.Media.SystemSounds.Beep.Play();
            dr = MessageBox.Show("Выйти из программы?", "Внимание", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1);

            if (dr == DialogResult.Yes)         //если подтверждаем выход из программы
            {
                if (!IsRun)
                {
                    Dispose();
                }

                else
                {
                    System.Media.SystemSounds.Beep.Play(); System.Media.SystemSounds.Beep.Play(); System.Media.SystemSounds.Beep.Play();
                    MessageBox.Show("Для начала остановите процесс измерений (нажмите кнопку <<Остановить измерения>> основной формы", "Внимание");
                    e.Cancel = true;
                }
            }

            else
            {
                e.Cancel = true;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            CultureInfo inf = new CultureInfo(System.Threading.Thread.CurrentThread.CurrentCulture.Name);
            System.Threading.Thread.CurrentThread.CurrentCulture = inf;
            inf.NumberFormat.NumberDecimalSeparator = ".";
            button5.Enabled = false;

            cl = new Calculate();

            buff = new StringBuilder(1000);

            toolTip1.SetToolTip(this.label20, "Допуск: от 15 до 25 0С");
            toolTip1.SetToolTip(this.textBox15, "Допуск: от 15 до 25 0С");

            toolTip1.SetToolTip(this.label18, "Допуск: от 50 до 80 %");
            toolTip1.SetToolTip(this.textBox14, "Допуск: от 50 до 80 %");

            toolTip1.SetToolTip(this.label15, "Допуск: от 720 до 780 мм рт. ст.");
            toolTip1.SetToolTip(this.textBox8, "Допуск: от 720 до 780 мм рт. ст.");

            toolTip1.SetToolTip(this.label1, "Допуск: от 215 до 225 В");
            toolTip1.SetToolTip(this.textBox5, "Допуск: от 215 до 225 В");

            toolTip1.SetToolTip(this.label11, "Допуск: от 49.5 до 50.5 Гц");
            toolTip1.SetToolTip(this.textBox7, "Допуск: от 49.5 до 50.5 Гц");

            StreamReader sr = new StreamReader("path.txt");
            textBox6.Text = sr.ReadLine();
            sr.Close();

            sr = new StreamReader("adress_zva.txt");
            textBox1.Text = sr.ReadLine();
            sr.Close();

            sr = new StreamReader("adress_cnt.txt");
            textBox2.Text = sr.ReadLine();
            sr.Close();
        }

        //private void button10_Click(object sender, EventArgs e)
        //{
        //    if (!button10.Enabled)
        //    {
        //        return;
        //    }

        //    button5.PerformClick();

        //    button5.Enabled = false;
        //    button6.Enabled = false;
        //    button10.Enabled = false;

        //    groupBox2.Enabled = false;
        //    groupBox3.Enabled = false;
        //    groupBox5.Enabled = false;
        //    ///groupBox6.Enabled = false;
        //    //groupBox7.Enabled = false;
        //    groupBox8.Enabled = false;

        //    button1.Enabled = false;
        //    button7.Enabled = false;

        //    int count_port = 0;

        //    int[] number_port = new int[4];

        //    int k = 0;

        //    if (checkBox18.Checked)
        //    {
        //        count_port++;
        //        number_port[k] = 1;
        //        k++;
        //    }

        //    if (checkBox19.Checked)
        //    {
        //        count_port++;
        //        number_port[k] = 2;
        //        k++;
        //    }

        //    if (checkBox20.Checked)
        //    {
        //        count_port++;
        //        number_port[k] = 3;
        //        k++;
        //    }

        //    if (checkBox21.Checked)
        //    {
        //        count_port++;
        //        number_port[k] = 4;
        //    }

        //    DialogResult dr;


        //    try
        //    {
        //        if (IsConnectN5242)
        //        {
        //            double[] freqList = { 0.0003, 4, 8 };

        //            int[] dopusk1 = { -18, -14 };
        //            int[] dopusk2 = { -17, -13 };

        //            VAC.Write("SOUR1:POW -10");

        //            string str = CreateFolder(textBox6.Text);

        //            if (str == String.Empty)
        //            {
        //                System.Media.SystemSounds.Beep.Play(); System.Media.SystemSounds.Beep.Play(); System.Media.SystemSounds.Beep.Play();
        //                MessageBox.Show("Измените расположение файла поверки. Затем начните поверку заново", "Внимание");
        //                return;
        //            }

        //            string fileName = str + "\\" + String.Format("koeff_refl_{0}.txt", DateTime.Now.ToShortTimeString().Replace(":", " "));

        //            StreamWriter sw = new StreamWriter(FileStream.Null);

        //            try
        //            {
        //                sw = new StreamWriter(fileName);
        //            }

        //            catch
        //            {
        //                System.Media.SystemSounds.Beep.Play(); System.Media.SystemSounds.Beep.Play(); System.Media.SystemSounds.Beep.Play();
        //                MessageBox.Show("Отсутствуют права на запись файла по указанному пути, обратитесь к администратору", "Внимание", MessageBoxButtons.OK);
        //                return;
        //            }

        //            sw.WriteLine("Определение модуля коэффициента отражения порта в режимах источника и приемника сигнала");
        //            sw.WriteLine();
        //            sw.WriteLine(String.Format("{0,20}", "f1, ГГц") + "\t" + String.Format("{0,20}", "f2, ГГц, дБм") + "\t" + String.Format("{0,20}", "S11 (источник), дБ") + "\t" + String.Format("{0,20}", "Допуск. (макс), дБ") + "\t" + String.Format("{0,20}", "Соответ.?") + "\t" + String.Format("{0,20}", "S11 (приемник), дБ") + "\t" + String.Format("{0,20}", "Допуск. (макс), дБ") + "\t" + String.Format("{0,20}", "Соответ.?"));

        //            sw.Flush();

        //            double[] data = null;

        //            double[] data1 = null;

        //            double[] data_X = null;

        //            IsRun = true;

        //            System.Media.SystemSounds.Beep.Play(); System.Media.SystemSounds.Beep.Play(); System.Media.SystemSounds.Beep.Play();
        //            dr = MessageBox.Show(String.Format("Проведите калибровку ПОРТа_1 анализатора цепей (вместе с кабелем), после этого нажмите ОК"), "Внимание", MessageBoxButtons.OKCancel);

        //            if (dr == DialogResult.Cancel)
        //            {
        //                if (VAC != null)
        //                {
        //                    // Close the driver
        //                    VAC.Dispose();
        //                }

        //                return;
        //            }

        //            System.Media.SystemSounds.Beep.Play(); System.Media.SystemSounds.Beep.Play(); System.Media.SystemSounds.Beep.Play();
        //            dr = MessageBox.Show(String.Format("Проведите калибровку ПОРТа_2 анализатора цепей (вместе с кабелем), после этого нажмите ОК"), "Внимание", MessageBoxButtons.OKCancel);

        //            if (dr == DialogResult.Cancel)
        //            {
        //                if (VAC != null)
        //                {
        //                    // Close the driver
        //                    VAC.Dispose();
        //                }

        //                return;
        //            }

        //            for (int ii = 0; ii < count_port; ii++)
        //            {
        //                sw.WriteLine();
        //                sw.WriteLine("Порт {0}", number_port[ii]);
        //                System.Media.SystemSounds.Beep.Play(); System.Media.SystemSounds.Beep.Play(); System.Media.SystemSounds.Beep.Play();

        //                if (number_port[ii] != 1)
        //                {
        //                    dr = MessageBox.Show(String.Format("Подключите к ПОРТу_{0} анализатора цепей (без кабеля) ПОРТ_1 (с кабелем), после этого нажмите ОК", number_port[ii]), "Внимание", MessageBoxButtons.OKCancel);
        //                }
        //                else
        //                {
        //                    dr = MessageBox.Show(String.Format("Подключите к ПОРТу_1 анализатора цепей (без кабеля) ПОРТ_2 (с кабелем), после этого нажмите ОК"), "Внимание", MessageBoxButtons.OKCancel);
        //                }

        //                if (dr == DialogResult.Cancel)
        //                {
        //                    if (VAC != null)
        //                    {
        //                        // Close the driver
        //                        VAC.Dispose();
        //                    }

        //                    return;
        //                }

        //                Thread.Sleep(500);

        //                if (number_port[ii] != 1)
        //                {
        //            //        Ch1S22.Create(1, 1); // Define Sii measurement
        //                }
        //                else
        //                {
        //             //       Ch1S22.Create(2, 2); // Define Sii measurement
        //                }

        //           //     Ch1S11.Create(number_port[ii], number_port[ii]); // Define Sii measurement                       

        //               // Thread.Sleep((int)(1.1 * Convert.ToDouble(temp.Query("SENS:SWE:TIME?").Replace("\n", "")) * 1000));

        //            //    Ch1S22.Trace.AutoScale();

        //             //   Thread.Sleep((int)(1.1 * Convert.ToDouble(temp.Query("SENS:SWE:TIME?").Replace("\n", "")) * 1000));
        //                Thread.Sleep(500);

        //            //    data = Ch1S22.FetchFormatted();

        //                if (number_port[ii] != 1)
        //                {
        //                    //Ch1S11.Create(number_port[ii], 1); // Define Sii measurement
        //                }
        //                else
        //                {
        //                //    Ch1S11.Create(number_port[ii], 2); // Define Sii measurement
        //                }

        //           //     Thread.Sleep((int)(1.1 * Convert.ToDouble(temp.Query("SENS:SWE:TIME?").Replace("\n", "")) * 1000));
        //                Thread.Sleep(500);

        //          //      Ch1S22.Trace.AutoScale();
        //                Thread.Sleep(500);
        //          //      Thread.Sleep((int)(1.1 * Convert.ToDouble(temp.Query("SENS:SWE:TIME?").Replace("\n", "")) * 1000));
        //                Thread.Sleep(500);

        //          //      data1 = Ch1S22.FetchFormatted();

        //         //       data_X = Ch1S22.FetchX();

        //                int number = 0;

        //                double max = -100;
        //                double max1 = -100;

        //                string results = "Нет";
        //                string results1 = "Нет";

        //                for (int ll = 1; ll < freqList.Length; ll++)
        //                {
        //                    max = -100;
        //                    max1 = -100;

        //                    for (int l = number + 1; l < data_X.Length; l++)
        //                    {
        //                        if (data_X[l] >= freqList[ll] * Math.Pow(10, 9))
        //                        {
        //                            for (int j = number; j <= l; j++)
        //                            {
        //                                if (max < data[j])
        //                                {
        //                                    max = data[j];
        //                                }

        //                                if (max1 < data1[j])
        //                                {
        //                                    max1 = data1[j];
        //                                }
        //                            }

        //                            number = l;

        //                            if (max < dopusk1[ll - 1])
        //                            {
        //                                results = "Да";
        //                            }
        //                            else
        //                            {
        //                                results = "Нет";
        //                            }

        //                            if (max1 < dopusk2[ll - 1])
        //                            {
        //                                results1 = "Да";
        //                            }
        //                            else
        //                            {
        //                                results1 = "Нет";
        //                            }

        //                            break;
        //                        }                            
        //                    }

        //                    sw.WriteLine(String.Format("{0,20}", freqList[ll - 1].ToString()) + "\t" + String.Format("{0,20}", freqList[ll].ToString()) + "\t" + String.Format("{0,20}", (max).ToString()) + "\t" + String.Format("{0,20}", (dopusk1[ll - 1]).ToString()) + "\t" + String.Format("{0,20}", results) + "\t" + String.Format("{0,20}", (max1).ToString()) + "\t" + String.Format("{0,20}", (dopusk2[ll - 1]).ToString()) + "\t" + String.Format("{0,20}", results1));             
        //                }

        //                Thread.Sleep(500);
        //                sw.Flush();

        //                if (ii == count_port - 1)
        //                {
        //                    sw.Close();
        //                    System.Media.SystemSounds.Beep.Play(); System.Media.SystemSounds.Beep.Play(); System.Media.SystemSounds.Beep.Play();
        //                    dr = MessageBox.Show("Процедура поверки по КСВН закончена. Просмотреть результаты?", "Внимание", MessageBoxButtons.OKCancel);

        //                    if (dr == DialogResult.OK)
        //                    {
        //                        Process.Start(fileName);
        //                    }
        //                }

        //            }
        //        }
        //        else
        //        {
        //            System.Media.SystemSounds.Beep.Play(); System.Media.SystemSounds.Beep.Play(); System.Media.SystemSounds.Beep.Play();
        //            MessageBox.Show("Проверьте подключение анализатора цепей к компьютеру", "Внимание", MessageBoxButtons.OK);
        //        }
        //    }

        //    catch (Exception ex)
        //    {
        //        System.Media.SystemSounds.Beep.Play(); System.Media.SystemSounds.Beep.Play(); System.Media.SystemSounds.Beep.Play();
        //        MessageBox.Show("Нарушение инструкции измерений, повторите попытку поверки по КСВН еще раз", "Внимание");
        //    }

        //    finally
        //    {

        //        button5.Enabled = true;
        //        button6.Enabled = true;

        //        button10.Enabled = true;

        //        groupBox2.Enabled = true;
        //        groupBox3.Enabled = true;
        //        groupBox5.Enabled = true;
        //        //groupBox6.Enabled = true;
        //        //groupBox7.Enabled = true;
        //        groupBox8.Enabled = true;

        //        button1.Enabled = true;
        //        button7.Enabled = true;

        //        button9.Enabled = true;

        //        IsRun = false;
        //    }
        //}

        private void textBox15_KeyPress(object sender, KeyPressEventArgs e)
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

        private void textBox15_MouseLeave(object sender, EventArgs e)
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

            if (((TextBox)sender).Text[0] == ',' || ((TextBox)sender).Text[0] == '.')
            {
                ((TextBox)sender).Text = "0" + ((TextBox)sender).Text;
            }

            if (((TextBox)sender).Text.IndexOf(".") != -1)
            {
                if (((TextBox)sender).Text.Length == ((TextBox)sender).Text.IndexOf(".") + 1)
                {
                    ((TextBox)sender).Text += "0";
                }
            }

            if (((TextBox)sender).Text.IndexOf(",") != -1)
            {
                if (((TextBox)sender).Text.Length == ((TextBox)sender).Text.IndexOf(",") + 1)
                {
                    ((TextBox)sender).Text += "0";
                }
            }

            if (((TextBox)sender).Text.IndexOf(".") != -1 || ((TextBox)sender).Text.IndexOf(",") != -1)
            {
                while (((TextBox)sender).Text[((TextBox)sender).Text.Length - 1] == '0' && (((TextBox)sender).Text[((TextBox)sender).Text.Length - 2] != '.' && ((TextBox)sender).Text[((TextBox)sender).Text.Length - 2] != ','))
                {
                    ((TextBox)sender).Text = ((TextBox)sender).Text.Remove(((TextBox)sender).Text.Length - 1, 1);
                }
            }
        }

        private void textBox15_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if (Convert.ToDouble(textBox15.Text, provider) > 25 || Convert.ToDouble(textBox15.Text, provider) < 15)
                {
                    textBox15.BackColor = Color.Red;
                }
                else
                {
                    textBox15.BackColor = Color.White;
                }
            }
            catch { };
        }

        private void textBox14_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if (Convert.ToDouble(textBox14.Text, provider) > 80 || Convert.ToDouble(textBox14.Text, provider) < 50)
                {
                    textBox14.BackColor = Color.Red;
                }
                else
                {
                    textBox14.BackColor = Color.White;
                }
            }
            catch { };
        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if (Convert.ToDouble(textBox8.Text, provider) > 780 || Convert.ToDouble(textBox8.Text, provider) < 720)
                {
                    textBox8.BackColor = Color.Red;
                }
                else
                {
                    textBox8.BackColor = Color.White;
                }
            }
            catch { };
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if (Convert.ToDouble(textBox5.Text, provider) > 225 || Convert.ToDouble(textBox5.Text, provider) < 215)
                {
                    textBox5.BackColor = Color.Red;
                }
                else
                {
                    textBox5.BackColor = Color.White;
                }
            }
            catch { };
        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if (Convert.ToDouble(textBox7.Text, provider) > 50.5 || Convert.ToDouble(textBox7.Text, provider) < 49.5)
                {
                    textBox7.BackColor = Color.Red;
                }
                else
                {
                    textBox7.BackColor = Color.White;
                }
            }
            catch { };
        }

        private void comboBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = true;
        }

        private void button14_Click(object sender, EventArgs e)
        {
            date = Convert.ToDateTime(dateTimePicker1.Value);

            if (fm == null || fm.IsDisposed)
            {
                fm = new General();
                fm.Show();
            }
            else
            {
                fm.Activate();
            }

            button8.Enabled = true;
            button5.Enabled = true;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            if (!button8.Enabled)
            {
                return;
            }

            if (!IsConnectN5242)
            {
                System.Media.SystemSounds.Beep.Play(); System.Media.SystemSounds.Beep.Play(); System.Media.SystemSounds.Beep.Play();
                MessageBox.Show("Проверьте физическое и программное анализатора цепей к компьютеру", "Внимание", MessageBoxButtons.OK);
                return;
            }

            try
            {

                DialogResult dr;

                if (Convert.ToDouble(textBox15.Text, provider) > 25 || Convert.ToDouble(textBox15.Text, provider) < 15)
                {
                    textBox15.BackColor = Color.Red;
                    dr = MessageBox.Show(String.Format("Не соответствует условие поверки - {0}. Продолжить поверку?", label20.Text), "Внимание", MessageBoxButtons.OKCancel);

                    if (dr == DialogResult.Cancel)
                    {
                        return;
                    }
                }
                else
                {
                    textBox15.BackColor = Color.White;
                }


                if (Convert.ToDouble(textBox14.Text, provider) > 80 || Convert.ToDouble(textBox14.Text, provider) < 50)
                {
                    textBox14.BackColor = Color.Red;
                    dr = MessageBox.Show(String.Format("Не соответствует условие поверки - {0}. Продолжить поверку?", label18.Text), "Внимание", MessageBoxButtons.OKCancel);

                    if (dr == DialogResult.Cancel)
                    {
                        return;
                    }
                }
                else
                {
                    textBox14.BackColor = Color.White;
                }


                if (Convert.ToDouble(textBox8.Text, provider) > 780 || Convert.ToDouble(textBox8.Text, provider) < 720)
                {
                    textBox8.BackColor = Color.Red;
                    dr = MessageBox.Show(String.Format("Не соответствует условие поверки - {0}. Продолжить поверку?", label15.Text), "Внимание", MessageBoxButtons.OKCancel);

                    if (dr == DialogResult.Cancel)
                    {
                        return;
                    }
                }
                else
                {
                    textBox8.BackColor = Color.White;
                }

                if (Convert.ToDouble(textBox5.Text, provider) > 225 || Convert.ToDouble(textBox5.Text, provider) < 215)
                {
                    textBox5.BackColor = Color.Red;
                    dr = MessageBox.Show(String.Format("Не соответствует условие поверки - {0}. Продолжить поверку?", label1.Text), "Внимание", MessageBoxButtons.OKCancel);

                    if (dr == DialogResult.Cancel)
                    {
                        return;
                    }
                }
                else
                {
                    textBox5.BackColor = Color.White;
                }

                if (Convert.ToDouble(textBox7.Text, provider) > 50.5 || Convert.ToDouble(textBox7.Text, provider) < 49.5)
                {
                    textBox7.BackColor = Color.Red;
                    dr = MessageBox.Show(String.Format("Не соответствует условие поверки - {0}. Продолжить поверку?", label11.Text), "Внимание", MessageBoxButtons.OKCancel);

                    if (dr == DialogResult.Cancel)
                    {
                        return;
                    }
                }
                else
                {
                    textBox7.BackColor = Color.White;
                }

                path = CreateFolder(textBox6.Text);

                if (path == String.Empty)
                {
                    System.Media.SystemSounds.Beep.Play(); System.Media.SystemSounds.Beep.Play(); System.Media.SystemSounds.Beep.Play();
                    MessageBox.Show("Измените расположение файла поверки. Затем начните поверку заново", "Внимание");
                    return;
                }

                string fileName = Form1.path + "\\" + String.Format("general_information_{0}.txt", date.ToShortTimeString().Replace(":", " "));

                StreamWriter sw = new StreamWriter(FileStream.Null);

                try
                {
                    sw = new StreamWriter(fileName);
                }

                catch
                {
                    System.Media.SystemSounds.Beep.Play(); System.Media.SystemSounds.Beep.Play(); System.Media.SystemSounds.Beep.Play();
                    MessageBox.Show("Отсутствуют права на запись файла по указанному пути, обратитесь к администратору", "Внимание", MessageBoxButtons.OK);
                    sw.Close();
                    return;
                }

                sw.WriteLine(General.predrp);
                sw.WriteLine();

                sw.WriteLine(General.addres);
                sw.WriteLine();

                sw.WriteLine("Прибор поверен в соответствии с Методикой поверки: {0}", General.metodik_pover);
                sw.WriteLine();

                sw.WriteLine("Эталоны и средства измерений, применяемые при поверке");
                sw.WriteLine();

                sw.WriteLine(General.etalon_name1 + " годен до " + General.etalon_data1 + " г.");
                sw.WriteLine(General.etalon_name2 + " годен до " + General.etalon_data2 + " г.");
                sw.WriteLine(General.etalon_name3 + " годен до " + General.etalon_data3 + " г.");
                sw.WriteLine(General.etalon_name4 + " годен до " + General.etalon_data4 + " г.");
                sw.WriteLine(General.etalon_name5 + " годен до " + General.etalon_data5 + " г.");

                sw.WriteLine();

                sw.WriteLine("ZVA24");
                sw.WriteLine();

                string[] temp = path.Split('\\');

                string number = temp[temp.Length - 2].Substring(temp[temp.Length - 2].Length - 6, 6);

                sw.WriteLine(String.Format("Зав №: {0}", number));
                sw.WriteLine();

                sw.WriteLine(String.Format("Дата поверки: {0}", dateTimePicker1.Text));
                sw.WriteLine();

                sw.WriteLine(String.Format("Проверку провел: {0}", textBox17.Text));
                sw.WriteLine();

                sw.WriteLine(String.Format("Принадлежит: {0}", textBox18.Text));
                sw.WriteLine();

                string type_pov;
                if (checkBox10.Checked)
                {
                    type_pov = "первичная";
                }
                else
                {
                    type_pov = "периодическая";
                }
                sw.WriteLine(String.Format("Вид поверки: {0}", type_pov));
                sw.WriteLine();

                sw.WriteLine(String.Format("Номер протокола: {0}", General.number_prot));
                sw.WriteLine();

                sw.WriteLine(groupBox10.Text);
                sw.WriteLine();

                sw.WriteLine(label20.Text + ": " + textBox15.Text);
                sw.WriteLine(label18.Text + ": " + textBox14.Text);
                sw.WriteLine(label15.Text + ": " + textBox8.Text);
                sw.WriteLine(label1.Text + ": " + textBox5.Text);
                sw.WriteLine(label11.Text + ": " + textBox7.Text);

                sw.WriteLine();
                sw.WriteLine("Результаты поверки");
                sw.WriteLine();

                sw.WriteLine(label28.Text);
                sw.WriteLine(comboBox1.Text);
                sw.WriteLine();

                sw.WriteLine(label27.Text);
                sw.WriteLine(comboBox3.Text);
                sw.WriteLine();

                sw.Close();

                dr = MessageBox.Show("Просмотреть файл?", "Внимание", MessageBoxButtons.OKCancel);

                if (dr == DialogResult.OK)
                {
                    Process.Start(fileName);
                }
            }

            catch
            {
                System.Media.SystemSounds.Beep.Play(); System.Media.SystemSounds.Beep.Play(); System.Media.SystemSounds.Beep.Play();
                MessageBox.Show("Проверьте физическое и программное подключение анализатора цепей к компьютеру", "Внимание", MessageBoxButtons.OK);
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {

        }

        private void button13_Click(object sender, EventArgs e)
        {
            IsRun = false;
            EndCalculate();
        }

        private void button16_Click(object sender, EventArgs e)
        {
            if (!button16.Enabled)
            {
                return;
            }

            if (!IsConnectN5242)
            {
                System.Media.SystemSounds.Beep.Play(); System.Media.SystemSounds.Beep.Play(); System.Media.SystemSounds.Beep.Play();
                MessageBox.Show("Проверьте физическое и программное подключение анализатора цепей к компьютеру", "Внимание", MessageBoxButtons.OK);
                return;
            }

            int count_port = 0;

            int[] number_port = new int[4];

            int k = 0;

            if (checkBox36.Checked)
            {
                count_port++;
                number_port[k] = 1;
                k++;
            }

            if (checkBox35.Checked)
            {
                count_port++;
                number_port[k] = 2;
                k++;
            }
            if (checkBox13.Checked)
            {
                count_port++;
                number_port[k] = 3;
                k++;
            }
            if (checkBox12.Checked)
            {
                count_port++;
                number_port[k] = 4;
                k++;
            }

            string temp = String.Empty;

            for (int i = 0; i < 4; i++)
            {
                if (number_port[i] != 0)
                {
                    if (temp == String.Empty)
                    {
                        temp += $"{number_port[i]}";
                    }
                    else
                    {
                        temp += $", {number_port[i]}";
                    }
                }
            }
            VAC.Timeout = 20000;

            VAC.Write("*RST");

            Thread.Sleep(200);
            VAC.Write(@"SWE:POIN 1000");
            Thread.Sleep(200);
            VAC.Write(@"SENS1:BWID 1000");
            Thread.Sleep(200);
            VAC.Write(@"INIT:CONT ON");
            Thread.Sleep(200);

            DialogResult dr = MessageBox.Show($"Переведите анализатор цепей в режим Go To Local, проведите полную однопортовую калибровку портов {temp} без кабеля, после этого нажмите ОК", "Внимание", MessageBoxButtons.OKCancel);

            if (dr == DialogResult.Cancel)
            {
                return;
            }

            try
            {
                path = CreateFolder(textBox6.Text);

                if (path == String.Empty)
                {
                    System.Media.SystemSounds.Beep.Play(); System.Media.SystemSounds.Beep.Play(); System.Media.SystemSounds.Beep.Play();
                    MessageBox.Show("Измените расположение файла поверки. Затем начните поверку заново", "Внимание");
                    return;
                }

                string fileName = path + "\\" + String.Format("refl_{0}.txt", Form1.date.ToShortTimeString().Replace(":", " "));
                StreamWriter sw = new StreamWriter(FileStream.Null);

                try
                {
                    sw = new StreamWriter(fileName);
                }

                catch
                {
                    System.Media.SystemSounds.Beep.Play(); System.Media.SystemSounds.Beep.Play(); System.Media.SystemSounds.Beep.Play();
                    MessageBox.Show("Отсутствуют права на запись файла по указанному пути, обратитесь к администратору", "Внимание", MessageBoxButtons.OK);
                    return;
                }

                IsRun = true;
                button2.Enabled = false;
                button5.Enabled = false;
                button6.Enabled = false;

                groupBox5.Enabled = false;
                groupBox2.Enabled = false;


                progressBar1.Visible = false;

                sw.WriteLine("Определение абсолютной погрешности измерений модуля и фазы коэффициента отражения");
                sw.WriteLine();

                string[] str_data;

                double[] data;
                double[] data_X;

                string result;

                List<double> freq = new List<double>();                
                List<double> swr_abs = new List<double>();
                List<double> swr_phase = new List<double>();
                List<double> swr_dopusk = new List<double>();
                List<double> swr_phase_dopusk = new List<double>();

                StreamReader sr = new StreamReader("swr_et_abs.txt");

                sr.ReadLine();

                while (!sr.EndOfStream)
                {
                    string[] temp1 = sr.ReadLine().Trim().Split('\t');

                    freq.Add(Convert.ToDouble(temp1[0].Trim().Replace(",", "."), provider));
                    swr_abs.Add(Convert.ToDouble(temp1[1].Trim().Replace(",", "."), provider));                   
                    swr_dopusk.Add(Convert.ToDouble(temp1[2].Trim().Replace(",", "."), provider));
                }

                sr.Close();

                sr = new StreamReader("swr_et_phase.txt");

                while (!sr.EndOfStream)
                {
                    string[] temp1 = sr.ReadLine().Trim().Split('\t');                    
                    swr_phase.Add(Convert.ToDouble(temp1[1].Trim().Replace(",", "."), provider));                  
                    swr_phase_dopusk.Add(Convert.ToDouble(temp1[2].Trim().Replace(",", "."), provider));
                }

                sr.Close();

                for (int i = 0; i < count_port; i++)
                {
                    dr = MessageBox.Show($"Подключите меру коэффициента отражения к порту {number_port[i]} анализатора, после этого нажмите ОК", "Внимание", MessageBoxButtons.OKCancel);

                    if (dr == DialogResult.Cancel)
                    {
                        button2.Enabled = true;
                        button5.Enabled = true;
                        button6.Enabled = true;

                        groupBox5.Enabled = true;
                        groupBox2.Enabled = true;

                        progressBar1.Visible = false;
                        IsRun = false;

                        dr = MessageBox.Show("Процедура поверки по коэффициенту отражения закончена. Просмотреть результаты?", "Внимание", MessageBoxButtons.OKCancel);

                        if (dr == DialogResult.OK)
                        {
                            Process.Start(fileName);
                        }

                        return;
                    }

                    sw.WriteLine();
                    sw.WriteLine("Порт {0}", number_port[i]);

                    switch (number_port[i])
                    {
                        case (1):
                            VAC.Write(@"CALC1:PAR:MEAS 'Trc1', 'S11'");
                            break;

                        case (2):
                            VAC.Write(@"CALC1:PAR:MEAS 'Trc1', 'S22'");
                            break;

                        case (3):
                            VAC.Write(@"CALC1:PAR:MEAS 'Trc1', 'S33'");
                            break;

                        case (4):
                            VAC.Write(@"CALC1:PAR:MEAS 'Trc1', 'S44'");
                            break;
                    }

                    sw.WriteLine("Модуль коэффициента отражения");
                    sw.WriteLine(String.Format("{0,20}", "f_ген, ГГц") + "\t" + String.Format("{0,20}", "|Sii|, дБ") + "\t" + String.Format("{0,20}", "|Sэii|, дБ") + "\t" + String.Format("{0,20}", "Погрешн.изм. |Sii|, дБ") + "\t" + String.Format("{0,20}", "Доп.знач., (±), дБ") + "\t" + String.Format("{0,20}", "Соответ.?"));
                    VAC.Write(@"CALC1:FORM MLOG");
                    Thread.Sleep((int)(1.1 * Convert.ToDouble(VAC.Query("SWE:TIME?").Replace("\n", ""), provider) * 1000));

                    str_data = VAC.Query("CALC1:DATA? FDAT").Split(',');
                    data = new double[str_data.Length];
                    data_X = new double[str_data.Length];

                    for (int iii = 0; iii < str_data.Length; iii++)
                    {
                        data[iii] = Convert.ToDouble(str_data[iii], provider);
                        data_X[iii] = 10 * Math.Pow(10, 6) + (iii) * (24 * Math.Pow(10, 9) - 10 * Math.Pow(10, 6)) / (1000.0 - 1);
                    }

                    for (int iii = 0; iii < freq.Count; iii++)
                    {
                        if (Math.Abs(GetValueTrue(data, data_X, FindIndex(freq[iii] * Math.Pow(10, 9), data_X), freq[iii]) - swr_abs[iii]) <= swr_dopusk[iii])
                        {
                            result = "Да";
                        }
                        else
                        {
                            result = "Нет";
                        }

                        sw.WriteLine((String.Format("{0,20}", (freq[iii]).ToString() + $"S{number_port[i]},{number_port[i]}") + "\t" + String.Format("{0,20}", GetValueTrue(data, data_X, FindIndex(freq[iii] * Math.Pow(10, 9), data_X), freq[iii]).ToString()) + "\t" + String.Format("{0,20}", swr_abs[iii]) + "\t" + String.Format("{0,20}", (Math.Round(GetValueTrue(data, data_X, FindIndex(freq[iii] * Math.Pow(10, 9), data_X), freq[iii]) - swr_abs[iii], 2)).ToString()) + "\t" + String.Format("{0,20}", Math.Round(swr_dopusk[iii], 3).ToString()) + "\t" + String.Format("{0,20}", result)).Replace('.', ','));
                    }

                    sw.Flush();

                    sw.WriteLine();
                    sw.WriteLine("Фаза коэффициента отражения");
                    sw.WriteLine(String.Format("{0,20}", "f_ген, ГГц") + "\t" + String.Format("{0,20}", "|φii|") + "\t" + String.Format("{0,20}", "|φэii|") + "\t" + String.Format("{0,20}", "Погрешн.изм., град.") + "\t" + String.Format("{0,20}", "Доп.знач., град. (±)") + "\t" + String.Format("{0,20}", "Соответ.?"));
                    VAC.Write(@"CALC1:FORM PHAS");
                    Thread.Sleep((int)(1.1 * Convert.ToDouble(VAC.Query("SWE:TIME?").Replace("\n", ""), provider) * 1000));

                    str_data = VAC.Query("CALC1:DATA? FDAT").Split(',');

                    data = new double[str_data.Length];
                    data_X = new double[str_data.Length];

                    for (int iii = 0; iii < str_data.Length; iii++)
                    {
                        data[iii] = Convert.ToDouble(str_data[iii], provider);
                        data_X[iii] = 10 * Math.Pow(10, 6) + (iii) * (24 * Math.Pow(10, 9) - 10 * Math.Pow(10, 6)) / (1000.0 - 1);
                    }

                    for (int iii = 0; iii < freq.Count; iii++)
                    {
                        if (Math.Abs(GetValueTrue(data, data_X, FindIndex(freq[iii] * Math.Pow(10, 9), data_X), freq[iii]) - swr_phase[iii]) <= swr_phase_dopusk[iii])
                        {
                            result = "Да";
                        }
                        else
                        {
                            result = "Нет";
                        }

                        sw.WriteLine((String.Format("{0,20}", (freq[iii]).ToString() + $"φ{number_port[i]},{number_port[i]}") + "\t" + String.Format("{0,20}", GetValueTrue(data, data_X, FindIndex(freq[iii] * Math.Pow(10, 9), data_X), freq[iii]).ToString()) + "\t" + String.Format("{0,20}", swr_phase[iii]) + "\t" + String.Format("{0,20}", (Math.Round(GetValueTrue(data, data_X, FindIndex(freq[iii] * Math.Pow(10, 9), data_X), freq[iii]) - swr_phase[iii], 2)).ToString()) + "\t" + String.Format("{0,20}", Math.Round(swr_phase_dopusk[iii], 3).ToString()) + "\t" + String.Format("{0,20}", result)).Replace('.', ','));
                    }

                    sw.Flush();
                }

                dr = MessageBox.Show("Процедура поверки по коэффициенту отражения закончена. Просмотреть результаты?", "Внимание", MessageBoxButtons.OKCancel);

                if (dr == DialogResult.OK)
                {
                    Process.Start(fileName);
                }

                button2.Enabled = true;
                button5.Enabled = true;
                button6.Enabled = true;

                groupBox5.Enabled = true;
                groupBox2.Enabled = true;


                progressBar1.Visible = false;
                IsRun = false;

            }

            catch (Exception ex)
            {
                System.Media.SystemSounds.Beep.Play(); System.Media.SystemSounds.Beep.Play(); System.Media.SystemSounds.Beep.Play();
                MessageBox.Show("Нарушение инструкции измерений, повторите попытку поверки коэффициенту отражения еще раз", "Внимание");
            }

        }


        private int FindIndex(double freq, double[] ar)
        {
            int index = 0;

            double freq_min = ar[0];
            double freq_max = ar[1];

            for (int i = 1; i < ar.Length; i++)
            {
                if (freq >= freq_min && freq <= freq_max)
                {
                    index = i;
                    break;
                }
                else
                {
                    freq_min = ar[i];
                    freq_max = ar[i + 1];
                }
            }

            return index;
        }

        private void button15_Click(object sender, EventArgs e)
        {
            date = Convert.ToDateTime(dateTimePicker1.Value);

            if (fm2 == null || fm2.IsDisposed)
            {
                fm2 = new Form3();
                fm2.Show();
            }
            else
            {
                fm2.Activate();
            }

            button16.Enabled = true;
        }

        private void button17_Click(object sender, EventArgs e)
        {
            date = Convert.ToDateTime(dateTimePicker1.Value);

            if (fm2 == null || fm2.IsDisposed)
            {
                fm2 = new Form3();
                fm2.Show();
            }
            else
            {
                fm2.Activate();
            }

            button16.Enabled = true;    
        }

        private void pictureBox18_Click(object sender, EventArgs e)
        {
            button16.PerformClick();
        }

        private double GetValueTrue(double[] data1, double[] data1_dx, int index, double freq)
        {
            index--;
            freq *= Math.Pow(10, 9);

            if (index == data1.Length - 1)
            {
                return data1[index];
            }

            double value = 0;

            value = data1[index] * (data1_dx[index + 1] - freq) / (data1_dx[index + 1] - data1_dx[index]) + data1[index + 1] * (freq - data1_dx[index]) / (data1_dx[index + 1] - data1_dx[index]);

            return value;
        }

        private void button18_Click(object sender, EventArgs e)
        {
            if (!button18.Enabled)
            {
                return;
            }

            if (!IsConnectN5242)
            {
                System.Media.SystemSounds.Beep.Play(); System.Media.SystemSounds.Beep.Play(); System.Media.SystemSounds.Beep.Play();
                MessageBox.Show("Проверьте физическое и программное подключение анализатора цепей к компьютеру", "Внимание", MessageBoxButtons.OK);
                return;
            }

            int count_port = 0;

            int[] number_port = new int[2];

            int k = 0;

            if (checkBox32.Checked)
            {
                count_port++;
                number_port[k] = 1;
                k++;
            }

            if (checkBox11.Checked)
            {
                count_port++;
                number_port[k] = 2;
                k++;
            }

            VAC.Timeout = 20000;
            VAC.Write("*RST");

            Thread.Sleep(200);
            VAC.Write(@"SWE:POIN 501");
            Thread.Sleep(200);
            VAC.Write(@"SENS1:BWID 200");
            Thread.Sleep(200);
            VAC.Write(@"INIT:CONT ON");
            Thread.Sleep(200);

            DialogResult dr = MessageBox.Show($"Переведите анализатор цепей в режим Go To Local, проведите полную двухпортовую калибровку поверяемых портов, после этого нажмите ОК", "Внимание", MessageBoxButtons.OKCancel);

            if (dr == DialogResult.Cancel)
            {
                return;
            }

            try
            {
                path = CreateFolder(textBox6.Text);

                if (path == String.Empty)
                {
                    System.Media.SystemSounds.Beep.Play(); System.Media.SystemSounds.Beep.Play(); System.Media.SystemSounds.Beep.Play();
                    MessageBox.Show("Измените расположение файла поверки. Затем начните поверку заново", "Внимание");
                    return;
                }

                string fileName = path + "\\" + String.Format("trans_{0}.txt", Form1.date.ToShortTimeString().Replace(":", " "));
                StreamWriter sw = new StreamWriter(FileStream.Null);

                try
                {
                    sw = new StreamWriter(fileName);
                }

                catch
                {
                    System.Media.SystemSounds.Beep.Play(); System.Media.SystemSounds.Beep.Play(); System.Media.SystemSounds.Beep.Play();
                    MessageBox.Show("Отсутствуют права на запись файла по указанному пути, обратитесь к администратору", "Внимание", MessageBoxButtons.OK);
                    return;
                }

                IsRun = true;
                button2.Enabled = false;
                button5.Enabled = false;
                button6.Enabled = false;

                groupBox5.Enabled = false;
                groupBox2.Enabled = false;

                progressBar1.Visible = false;

                sw.WriteLine("Определение абсолютной погрешности измерений модуля и фазы коэффициента передачи");
                sw.WriteLine();

                string[] str_data;

                double[] data;
                double[] data_X;

                double dopusk_abs;
                double dopusk_phase;
                string result;

                List<double> freq = new List<double>();
                List<double> trans_10_1_abs = new List<double>();
                List<double> trans_10_2_abs = new List<double>();
                List<double> trans_40_1_abs = new List<double>();
                List<double> trans_40_2_abs = new List<double>();
                List<double> trans_60_1_abs = new List<double>();
                List<double> trans_60_2_abs = new List<double>();

                List<double> trans_10_1_phase = new List<double>();
                List<double> trans_10_2_phase = new List<double>();
                List<double> trans_40_1_phase = new List<double>();
                List<double> trans_40_2_phase = new List<double>();
                List<double> trans_60_1_phase = new List<double>();
                List<double> trans_60_2_phase = new List<double>();

                List<double> abs = new List<double>();
                List<double> phase = new List<double>();

                StreamReader sr = new StreamReader("att_et_abs_12.txt");

                sr.ReadLine();

                while (!sr.EndOfStream)
                {
                    string[] temp1 = sr.ReadLine().Trim().Split('\t');

                    freq.Add(Convert.ToDouble(temp1[0].Trim().Replace(",", "."), provider));
                    trans_10_1_abs.Add(Convert.ToDouble(temp1[1].Trim().Replace(",", "."), provider));
                    trans_40_1_abs.Add(Convert.ToDouble(temp1[2].Trim().Replace(",", "."), provider));
                    trans_60_1_abs.Add(Convert.ToDouble(temp1[3].Trim().Replace(",", "."), provider));               
                }

                sr.Close();

                sr = new StreamReader("att_et_abs_21.txt");

                while (!sr.EndOfStream)
                {
                    string[] temp1 = sr.ReadLine().Trim().Split('\t');
                  
                    trans_10_2_abs.Add(Convert.ToDouble(temp1[1].Trim().Replace(",", "."), provider));
                    trans_40_2_abs.Add(Convert.ToDouble(temp1[2].Trim().Replace(",", "."), provider));
                    trans_60_2_abs.Add(Convert.ToDouble(temp1[3].Trim().Replace(",", "."), provider));
                }

                sr.Close();

                sr = new StreamReader("att_et_phase_12.txt");

                sr.ReadLine();

                while (!sr.EndOfStream)
                {
                    string[] temp1 = sr.ReadLine().Trim().Split('\t');
                  
                    trans_10_1_phase.Add(Convert.ToDouble(temp1[1].Trim().Replace(",", "."), provider));
                    trans_40_1_phase.Add(Convert.ToDouble(temp1[2].Trim().Replace(",", "."), provider));
                    trans_60_1_phase.Add(Convert.ToDouble(temp1[3].Trim().Replace(",", "."), provider));
                }

                sr.Close();

                sr = new StreamReader("att_et_phase_21.txt");

                while (!sr.EndOfStream)
                {
                    string[] temp1 = sr.ReadLine().Trim().Split('\t');
                   
                    trans_10_2_phase.Add(Convert.ToDouble(temp1[1].Trim().Replace(",", "."), provider));
                    trans_40_2_phase.Add(Convert.ToDouble(temp1[2].Trim().Replace(",", "."), provider));
                    trans_60_2_phase.Add(Convert.ToDouble(temp1[3].Trim().Replace(",", "."), provider));
                }

                sr.Close();

                string string_port = "";

                for (int i = 0; i < count_port; i++)
                {
                    sw.WriteLine();

                    if (i == 0)
                    {
                        string_port = "Порт 1-2";
                        sw.WriteLine("Порт 1-2");
                    }
                    else
                    {
                        string_port = "Порт 3-4";
                        sw.WriteLine("Порт 3-4");
                    }

                    for (int ij = 0; ij < 3; ij++)
                    {
                        if (ij == 0)
                        {
                            dr = MessageBox.Show($"Подключите аттенюатор 20 дБ между портами {string_port} анализатора, после этого нажмите ОК", "Внимание", MessageBoxButtons.OKCancel);
                        }
                        else if (ij == 1)
                        {
                            dr = MessageBox.Show($"Подключите аттенюатор 40 дБ между портами {string_port} анализатора, после этого нажмите ОК", "Внимание", MessageBoxButtons.OKCancel);
                        }
                        else
                        {
                            dr = MessageBox.Show($"Подключите аттенюатор 60 дБ между портами {string_port} анализатора, после этого нажмите ОК", "Внимание", MessageBoxButtons.OKCancel);
                        }

                        if (dr == DialogResult.Cancel)
                        {
                            button2.Enabled = true;
                            button5.Enabled = true;
                            button6.Enabled = true;

                            groupBox5.Enabled = true;
                            groupBox2.Enabled = true;


                            progressBar1.Visible = false;
                            IsRun = false;

                            dr = MessageBox.Show("Процедура поверки по коэффициенту передачи закончена. Просмотреть результаты?", "Внимание", MessageBoxButtons.OKCancel);

                            if (dr == DialogResult.OK)
                            {
                                Process.Start(fileName);
                            }

                            return;
                        }

                        for (int ii = 0; ii < 2; ii++)
                        {
                            if (i == 0)
                            {
                                if (ii == 0)
                                {
                                    VAC.Write(@"CALC1:PAR:MEAS 'Trc1', 'S12'");
                                }
                                else
                                {
                                    VAC.Write(@"CALC1:PAR:MEAS 'Trc1', 'S21'");
                                }
                            }
                            else
                            {
                                if (ii == 0)
                                {
                                    VAC.Write(@"CALC1:PAR:MEAS 'Trc1', 'S34'");
                                }
                                else
                                {
                                    VAC.Write(@"CALC1:PAR:MEAS 'Trc1', 'S43'");
                                }
                            }

                            if (ii == 0)
                            {
                                if (ij == 0)
                                {
                                    abs = trans_10_1_abs;
                                    phase = trans_10_1_phase;
                                }
                                else if (ij == 1)
                                {
                                    abs = trans_40_1_abs;
                                    phase = trans_40_1_phase;
                                }
                                else
                                {
                                    abs = trans_60_1_abs;
                                    phase = trans_60_1_phase;
                                }
                            }
                            else
                            {
                                if (ij == 0)
                                {
                                    abs = trans_10_2_abs;
                                    phase = trans_10_2_phase;
                                }
                                else if (ij == 1)
                                {
                                    abs = trans_40_2_abs;
                                    phase = trans_40_2_phase;
                                }
                                else
                                {
                                    abs = trans_60_2_abs;
                                    phase = trans_60_2_phase;
                                }
                            }

                            sw.WriteLine("Модуль коэффициента передачи");
                            sw.WriteLine(String.Format("{0,20}", "f_ген, ГГц") + "\t" + String.Format("{0,20}", "|Sij|, дБ") + "\t" + String.Format("{0,20}", "|Sэij|, дБ") + "\t" + String.Format("{0,20}", "Погрешн.изм. |Sij|, дБ") + "\t" + String.Format("{0,20}", "Доп.знач., (±), дБ") + "\t" + String.Format("{0,20}", "Соответ.?"));
                            VAC.Write(@"CALC1:FORM MLOG");
                            Thread.Sleep((int)(1.1 * Convert.ToDouble(VAC.Query("SWE:TIME?").Replace("\n", ""), provider) * 1000));

                            str_data = VAC.Query("CALC1:DATA? FDAT").Split(',');
                            data = new double[str_data.Length];
                            data_X = new double[str_data.Length];

                            for (int iii = 0; iii < str_data.Length; iii++)
                            {
                                data[iii] = Math.Abs(Convert.ToDouble(str_data[iii], provider));
                                data_X[iii] = 10 * Math.Pow(10, 6) + (iii) * (24 * Math.Pow(10, 9) - 10 * Math.Pow(10, 6)) / (501 - 1);
                            }

                            for (int iii = 0; iii < freq.Count; iii++)
                            {
                                if (ij == 0)
                                {
                                    if (freq[iii] < 0.05)
                                    {
                                        dopusk_abs = 1.0;
                                    }
                                    else if (freq[iii] < 0.4)
                                    {
                                        dopusk_abs = 0.2;
                                    }
                                    else
                                    {
                                        dopusk_abs = 0.1;
                                    }
                                }
                                else if (ij == 1)
                                {
                                    if (freq[iii] < 0.05)
                                    {
                                        dopusk_abs = 1.0;
                                    }
                                    else if (freq[iii] < 0.4)
                                    {
                                        dopusk_abs = 1.0;
                                    }
                                    else if(freq[iii] < 0.7)
                                    {
                                        dopusk_abs = 0.2;
                                    }
                                     else
                                    {
                                        dopusk_abs = 0.1;
                                    }
                                }
                                else
                                {
                                    if (freq[iii] < 0.05)
                                    {
                                        dopusk_abs = 1.0;
                                    }
                                    else if (freq[iii] < 0.4)
                                    {
                                        dopusk_abs = 1.0;
                                    }
                                    else if (freq[iii] < 0.7)
                                    {
                                        dopusk_abs = 1.0;
                                    }
                                    else
                                    {
                                        dopusk_abs = 0.2;
                                    }
                                }

                                if (Math.Abs(GetValueTrue(data, data_X, FindIndex(freq[iii] * Math.Pow(10, 9), data_X), freq[iii]) - abs[iii]) <= dopusk_abs)
                                {
                                    result = "Да";
                                }
                                else
                                {
                                    result = "Нет";
                                }

                                if (i == 0)
                                {
                                    if (ii == 0)
                                    {                                        
                                        sw.WriteLine((String.Format("{0,20}", (freq[iii]).ToString() + "S12") + "\t" + String.Format("{0,20}", GetValueTrue(data, data_X, FindIndex(freq[iii] * Math.Pow(10, 9), data_X), freq[iii]).ToString()) + "\t" + String.Format("{0,20}", abs[iii]) + "\t" + String.Format("{0,20}", (Math.Round(GetValueTrue(data, data_X, FindIndex(freq[iii] * Math.Pow(10, 9), data_X), freq[iii]) - abs[iii], 2)).ToString()) + "\t" + String.Format("{0,20}", Math.Round(dopusk_abs, 3).ToString()) + "\t" + String.Format("{0,20}", result)).Replace('.', ','));
                                    }
                                    else
                                    {
                                        sw.WriteLine((String.Format("{0,20}", (freq[iii]).ToString() + "S21") + "\t" + String.Format("{0,20}", GetValueTrue(data, data_X, FindIndex(freq[iii] * Math.Pow(10, 9), data_X), freq[iii]).ToString()) + "\t" + String.Format("{0,20}", abs[iii]) + "\t" + String.Format("{0,20}", (Math.Round(GetValueTrue(data, data_X, FindIndex(freq[iii] * Math.Pow(10, 9), data_X), freq[iii]) - abs[iii], 2)).ToString()) + "\t" + String.Format("{0,20}", Math.Round(dopusk_abs, 3).ToString()) + "\t" + String.Format("{0,20}", result)).Replace('.', ','));
                                    }
                                }
                                else
                                {
                                    if (ii == 0)
                                    {
                                        sw.WriteLine((String.Format("{0,20}", (freq[iii]).ToString() + "S34") + "\t" + String.Format("{0,20}", GetValueTrue(data, data_X, FindIndex(freq[iii] * Math.Pow(10, 9), data_X), freq[iii]).ToString()) + "\t" + String.Format("{0,20}", abs[iii]) + "\t" + String.Format("{0,20}", (Math.Round(GetValueTrue(data, data_X, FindIndex(freq[iii] * Math.Pow(10, 9), data_X), freq[iii]) - abs[iii], 2)).ToString()) + "\t" + String.Format("{0,20}", Math.Round(dopusk_abs, 3).ToString()) + "\t" + String.Format("{0,20}", result)).Replace('.', ','));
                                    }
                                    else
                                    {
                                        sw.WriteLine((String.Format("{0,20}", (freq[iii]).ToString() + "S43") + "\t" + String.Format("{0,20}", GetValueTrue(data, data_X, FindIndex(freq[iii] * Math.Pow(10, 9), data_X), freq[iii]).ToString()) + "\t" + String.Format("{0,20}", abs[iii]) + "\t" + String.Format("{0,20}", (Math.Round(GetValueTrue(data, data_X, FindIndex(freq[iii] * Math.Pow(10, 9), data_X), freq[iii]) - abs[iii], 2)).ToString()) + "\t" + String.Format("{0,20}", Math.Round(dopusk_abs, 3).ToString()) + "\t" + String.Format("{0,20}", result)).Replace('.', ','));
                                    }
                                }
                            }

                            sw.Flush();

                            sw.WriteLine();
                            sw.WriteLine("Фаза коэффициента передачи");
                            sw.WriteLine(String.Format("{0,20}", "f_ген, ГГц") + "\t" + String.Format("{0,20}", "|φii|") + "\t" + String.Format("{0,20}", "|φэii|") + "\t" + String.Format("{0,20}", "Погрешн.изм., град.") + "\t" + String.Format("{0,20}", "Доп.знач., град. (±)") + "\t" + String.Format("{0,20}", "Соответ.?"));
                            VAC.Write(@"CALC1:FORM PHAS");
                            Thread.Sleep((int)(1.1 * Convert.ToDouble(VAC.Query("SWE:TIME?").Replace("\n", ""), provider) * 1000));
                          //  Thread.Sleep(500);

                            str_data = VAC.Query("CALC1:DATA? FDAT").Split(',');

                            data = new double[str_data.Length];
                            data_X = new double[str_data.Length];

                            for (int iii = 0; iii < str_data.Length; iii++)
                            {
                                data[iii] = Convert.ToDouble(str_data[iii], provider);
                                data_X[iii] = 10 * Math.Pow(10, 6) + (iii) * (24 * Math.Pow(10, 9) - 10 * Math.Pow(10, 6)) / (501 - 1);
                            }

                            for (int iii = 0; iii < freq.Count; iii++)
                            {
                                if (ij == 0)
                                {
                                    if (freq[iii] < 0.05)
                                    {
                                        dopusk_phase = 6.0;
                                    }
                                    else if (freq[iii] < 0.4)
                                    {
                                        dopusk_phase = 2.0;
                                    }
                                    else
                                    {
                                        dopusk_phase = 1.0;
                                    }
                                }
                                else if (ij == 1)
                                {
                                    if (freq[iii] < 0.05)
                                    {
                                        dopusk_phase = 6.0;
                                    }
                                    else if (freq[iii] < 0.4)
                                    {
                                        dopusk_phase = 6.0;
                                    }
                                    else if (freq[iii] < 0.7)
                                    {
                                        dopusk_phase = 2.0;
                                    }
                                    else
                                    {
                                        dopusk_phase = 1.0;
                                    }
                                }
                                else
                                {
                                    if (freq[iii] < 0.05)
                                    {
                                        dopusk_phase = 6.0;
                                    }
                                    else if (freq[iii] < 0.4)
                                    {
                                        dopusk_phase = 6.0;
                                    }
                                    else if (freq[iii] < 0.7)
                                    {
                                        dopusk_phase = 6.0;
                                    }
                                    else
                                    {
                                        dopusk_phase = 2.0;
                                    }
                                }

                                if (Math.Abs(GetValueTrue(data, data_X, FindIndex(freq[iii] * Math.Pow(10, 9), data_X),freq[iii]) - phase[iii]) <= dopusk_phase)
                                {
                                    result = "Да";
                                }
                                else
                                {
                                    result = "Нет";
                                }

                                if (i == 0)
                                {
                                    if (ii == 0)
                                    {
                                        sw.WriteLine((String.Format("{0,20}", (freq[iii]).ToString() + "φ12") + "\t" + String.Format("{0,20}", GetValueTrue(data, data_X, FindIndex(freq[iii] * Math.Pow(10, 9), data_X), freq[iii]).ToString()) + "\t" + String.Format("{0,20}", phase[iii]) + "\t" + String.Format("{0,20}", (Math.Round(GetValueTrue(data, data_X, FindIndex(freq[iii] * Math.Pow(10, 9), data_X), freq[iii]) - phase[iii], 2)).ToString()) + "\t" + String.Format("{0,20}", Math.Round(dopusk_phase, 3).ToString()) + "\t" + String.Format("{0,20}", result)).Replace('.', ','));
                                    }
                                    else
                                    {
                                        sw.WriteLine((String.Format("{0,20}", (freq[iii]).ToString() + "φ21") + "\t" + String.Format("{0,20}", GetValueTrue(data, data_X, FindIndex(freq[iii] * Math.Pow(10, 9), data_X), freq[iii]).ToString()) + "\t" + String.Format("{0,20}", phase[iii]) + "\t" + String.Format("{0,20}", (Math.Round(GetValueTrue(data, data_X, FindIndex(freq[iii] * Math.Pow(10, 9), data_X), freq[iii]) - phase[iii], 2)).ToString()) + "\t" + String.Format("{0,20}", Math.Round(dopusk_phase, 3).ToString()) + "\t" + String.Format("{0,20}", result)).Replace('.', ','));
                                    }
                                }
                                else
                                {
                                    if (ii == 0)
                                    {
                                        sw.WriteLine((String.Format("{0,20}", (freq[iii]).ToString() + "φ34") + "\t" + String.Format("{0,20}", GetValueTrue(data, data_X, FindIndex(freq[iii] * Math.Pow(10, 9), data_X), freq[iii]).ToString()) + "\t" + String.Format("{0,20}", phase[iii]) + "\t" + String.Format("{0,20}", (Math.Round(GetValueTrue(data, data_X, FindIndex(freq[iii] * Math.Pow(10, 9), data_X), freq[iii]) - phase[iii], 2)).ToString()) + "\t" + String.Format("{0,20}", Math.Round(dopusk_phase, 3).ToString()) + "\t" + String.Format("{0,20}", result)).Replace('.', ','));
                                    }
                                    else
                                    {
                                        sw.WriteLine((String.Format("{0,20}", (freq[iii]).ToString() + "φ43") + "\t" + String.Format("{0,20}", GetValueTrue(data, data_X, FindIndex(freq[iii] * Math.Pow(10, 9), data_X), freq[iii]).ToString()) + "\t" + String.Format("{0,20}", phase[iii]) + "\t" + String.Format("{0,20}", (Math.Round(GetValueTrue(data, data_X, FindIndex(freq[iii] * Math.Pow(10, 9), data_X), freq[iii]) - phase[iii], 2)).ToString()) + "\t" + String.Format("{0,20}", Math.Round(dopusk_phase, 3).ToString()) + "\t" + String.Format("{0,20}", result)).Replace('.', ','));
                                    }
                                }
                            }

                            sw.Flush();
                        }
                    }
                }

                dr = MessageBox.Show("Процедура поверки по коэффициенту передачи закончена. Просмотреть результаты?", "Внимание", MessageBoxButtons.OKCancel);

                if (dr == DialogResult.OK)
                {
                    Process.Start(fileName);
                }

                button2.Enabled = true;
                button5.Enabled = true;
                button6.Enabled = true;

                groupBox5.Enabled = true;
                groupBox2.Enabled = true;
              

                progressBar1.Visible = false;
                IsRun = false;

            }

            catch (Exception ex)
            {
                System.Media.SystemSounds.Beep.Play(); System.Media.SystemSounds.Beep.Play(); System.Media.SystemSounds.Beep.Play();
                MessageBox.Show("Нарушение инструкции измерений, повторите попытку поверки коэффициенту передачи еще раз", "Внимание");
            }
        }

        private void pictureBox16_Click(object sender, EventArgs e)
        {
            button18.PerformClick();
        }

        private void button15_Click_1(object sender, EventArgs e)
        {
            date = Convert.ToDateTime(dateTimePicker1.Value);

            if (fm3 == null || fm3.IsDisposed)
            {
                fm3 = new Form2();
                fm3.Show();
            }
            else
            {
                fm3.Activate();
            }

            button19.Enabled = true;
        }

        private void button19_Click(object sender, EventArgs e)
        {
            date = Convert.ToDateTime(dateTimePicker1.Value);

            if (fm4 == null || fm4.IsDisposed)
            {
                fm4 = new Form4();
                fm4.Show();
            }
            else
            {
                fm4.Activate();
            }

            button18.Enabled = true;
        }
    }
}
