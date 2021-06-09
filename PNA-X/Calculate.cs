using NationalInstruments.VisaNS;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace PNA_X
{
    class Calculate
    {
        NumberFormatInfo provider = new NumberFormatInfo() { NumberDecimalSeparator = "." };
        Thread m_tread;                                    //вторичный поток, осуществляющий расчет

        DrawGraphicOnForm m_DrawGraphicOnForm;             //интерфейс, посредством которого осуществляется связь с основным потоком

        private MessageBasedSession CNT;
        private MessageBasedSession VAC;

        /// <summary> Запускаем расчет во вторичном потоке
        /// 
        /// </summary>
        /// <param name="p_DrawGraphicOnForm"></param>
        public void Start(DrawGraphicOnForm p_DrawGraphicOnForm)
        {
            m_DrawGraphicOnForm = p_DrawGraphicOnForm;
            m_tread = new Thread(new ThreadStart(Run));
            m_tread.Start();
        }

        /// <summary> Запускаем расчет во вторичном потоке
        /// 
        /// </summary>
        /// <param name="p_DrawGraphicOnForm"></param>
        public void Start1(DrawGraphicOnForm p_DrawGraphicOnForm)
        {
            m_DrawGraphicOnForm = p_DrawGraphicOnForm;
            m_tread = new Thread(new ThreadStart(Run1));
            m_tread.Start();
        }

        private void Run()
        {
            try
            {
                DialogResult dr;

                try
                {
                    // Create driver instance
                    CNT = (MessageBasedSession)ResourceManager.GetLocalManager().Open(Form1.addr_CNT);
                }
                catch
                {
                    MessageBox.Show("Проверьте физическое и программное подключение частотомера к компьютеру", "Внимание", MessageBoxButtons.OK);
                    m_DrawGraphicOnForm.EndCalculate();
                    return;
                }

                try
                {
                    // Create driver instance
                    VAC = (MessageBasedSession)ResourceManager.GetLocalManager().Open(Form1.addr_zva);
                }
                catch
                {
                    MessageBox.Show("Проверьте физическое и программное подключение анализатора цепей к компьютеру", "Внимание", MessageBoxButtons.OK);
                    m_DrawGraphicOnForm.EndCalculate();
                    return;
                }

                CNT.Timeout = 20000;
                VAC.Timeout = 20000;

                CNT.Write("*RST");
                VAC.Write("*RST");

                VAC.Write($"SENS:SWE:TYPE CW");
                CNT.Write($"SENS:ACQ:APER 300ms");

                VAC.Write("SOUR1:POW 3");

                string str1;
                string result;
                bool IsCanalA = false;
                bool IsCanalB = false;

                if (Form1.freqList_freq[0] <= 0.00001)
                {
                    MessageBox.Show("Некорректно задана измеряемая частота (за нижним пределом диапазона частот стенда)", "Внимание");
                    if (CNT != null)
                    {
                        // Close the driver
                        CNT.Dispose();
                    }
                    if (VAC != null)
                    {
                        // Close the driver
                        VAC.Dispose();
                    }
                    m_DrawGraphicOnForm.EndCalculate();
                    return;
                }

                if (Form1.freqList_freq[0] <= 0.3)
                {
     
                    IsCanalA = true;
                }
                else if (Form1.freqList_freq[0] <= 40)
                {
                   IsCanalB = true;
                }
                else
                {
                    MessageBox.Show("Некорректно задана измеряемая частота (за верхним пределом диапазона частот стенда)", "Внимание");
                    if (CNT != null)
                    {
                        // Close the driver
                        CNT.Dispose();
                    }
                    if (VAC != null)
                    {
                        // Close the driver
                        VAC.Dispose();
                    }
                    m_DrawGraphicOnForm.EndCalculate();
                    return;
                }

                string fileName = Form1.path + "\\" + String.Format("freq_{0}.txt", Form1.date.ToShortTimeString().Replace(":", " "));

                StreamWriter sw = new StreamWriter(FileStream.Null);

                try
                {
                    sw = new StreamWriter(fileName);
                }

                catch
                {
                    System.Media.SystemSounds.Beep.Play(); System.Media.SystemSounds.Beep.Play(); System.Media.SystemSounds.Beep.Play();
                    MessageBox.Show("Отсутствуют права на запись файла по указанному пути, обратитесь к администратору", "Внимание", MessageBoxButtons.OK);
                    if (CNT != null)
                    {
                        // Close the driver
                        CNT.Dispose();
                    }
                    if (VAC != null)
                    {
                        // Close the driver
                        VAC.Dispose();
                    }
                    m_DrawGraphicOnForm.EndCalculate();
                    return;
                }

                sw.WriteLine("Определение диапазона рабочих частот и относительной погрешности установки частоты источника выходного сигнала");
                sw.WriteLine();

                sw.WriteLine(String.Format("{0,20}", "f_уст, ГГц") + "\t" + String.Format("{0,20}", "f_изм, ГГц") + "\t" + String.Format("{0,20}", "Погрешн., отн.ед.") + "\t" + String.Format("{0,20}", "Доп.знач., отн.ед (±)") + "\t" + String.Format("{0,20}", "Соответ.?"));
                sw.Flush();
                Form1.IsRun = true;
                sw.WriteLine();


                for (int ii = 0; ii < Form1.count_port_freq; ii++)
                {
                    IsCanalA = false;
                    IsCanalB = false;

                    sw.WriteLine();
                    sw.WriteLine("Порт {0}", Form1.number_port_freq[ii]);

                    if (Form1.freqList_freq[0] <= 0.3)
                    {
                        dr = MessageBox.Show(String.Format("Подключите ПОРТ_{0} анализатора цепей к входу А частотомера, после этого нажмите ОК", Form1.number_port_freq[ii]), "Внимание", MessageBoxButtons.OKCancel);

                        IsCanalA = true;
                    }
                    else
                    {
                        dr = MessageBox.Show(String.Format("Подключите ПОРТ_{0} анализатора цепей к входу C частотомера, после этого нажмите ОК", Form1.number_port_freq[ii]), "Внимание", MessageBoxButtons.OKCancel);
                        IsCanalB = true;
                    }

                    if (dr == DialogResult.Cancel)
                    {
                        sw.Close();
                        Form1.IsRun = false;
                        if (CNT != null)
                        {
                            // Close the driver
                            CNT.Dispose();
                        }
                        if (VAC != null)
                        {
                            // Close the driver
                            VAC.Dispose();
                        }
                        dr = MessageBox.Show("Процедура поверки по частоте закончена. Просмотреть результаты?", "Внимание", MessageBoxButtons.OKCancel);

                        if (dr == DialogResult.OK)
                        {
                            Process.Start(fileName);
                        }
                        m_DrawGraphicOnForm.EndCalculate();
                        return;
                    }


                    for (int i = 0; i < Form1.freqList_freq.Length; i++)
                    {
                        if (!Form1.IsRun)
                        {
                            sw.Close();
                            Form1.IsRun = false;

                            if (CNT != null)
                            {
                                // Close the driver
                                CNT.Dispose();
                            }

                            if (VAC != null)
                            {
                                // Close the driver
                                VAC.Dispose();
                            }

                            dr = MessageBox.Show("Процедура поверки по частоте закончена. Просмотреть результаты?", "Внимание", MessageBoxButtons.OKCancel);

                            if (dr == DialogResult.OK)
                            {
                                Process.Start(fileName);
                            }
                            m_DrawGraphicOnForm.EndCalculate();
                            return;
                        }

                        if (i != 0)
                        {
                            if (Form1.freqList_freq[i] <= 0.3)
                            {

                            }

                            else if (Form1.freqList_freq[i] > 0.3 && Form1.freqList_freq[i] <= 40)
                            {
                                if (!IsCanalB)
                                {
                                    dr = MessageBox.Show(String.Format("Подключите ПОРТ_{0} анализатора цепей к входу C частотомера, после этого нажмите ОК", Form1.number_port_freq[ii]), "Внимание", MessageBoxButtons.OKCancel);

                                    IsCanalB = true;

                                    if (dr == DialogResult.Cancel)
                                    {
                                        if (CNT != null)
                                        {
                                            // Close the driver
                                            CNT.Dispose();
                                        }

                                        if (VAC != null)
                                        {
                                            // Close the driver
                                            VAC.Dispose();
                                        }


                                        sw.Close();
                                        Form1.IsRun = false;

                                        dr = MessageBox.Show("Процедура поверки по частоте закончена. Просмотреть результаты?", "Внимание", MessageBoxButtons.OKCancel);

                                        if (dr == DialogResult.OK)
                                        {
                                            Process.Start(fileName);
                                        }
                                        m_DrawGraphicOnForm.EndCalculate();
                                        return;
                                    }
                                }
                            }
                            else
                            {
                                MessageBox.Show("Некорректно задана измеряемая частота или измеряемая мощность (за диапазоном частот стенда)", "Внимание");


                                sw.Close();
                                Form1.IsRun = false;

                                dr = MessageBox.Show("Процедура поверки по частоте закончена. Просмотреть результаты?", "Внимание", MessageBoxButtons.OKCancel);

                                if (dr == DialogResult.OK)
                                {
                                    Process.Start(fileName);
                                }

                                if (CNT != null)
                                {
                                    // Close the driver
                                    CNT.Dispose();
                                }

                                if (VAC != null)
                                {
                                    // Close the driver
                                    VAC.Dispose();
                                }
                                m_DrawGraphicOnForm.EndCalculate();
                                return;
                            }
                        }

                        switch (Form1.number_port_freq[ii])
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


                        VAC.Write($"SENS:FREQ {Form1.freqList_freq[i] * Math.Pow(10, 9)}");

                        Thread.Sleep(100);

                        if (Form1.freqList_freq[i] <= 0.3)
                        {
                            str1 = CNT.Query("MEAS:FREQ? (@1)n").Replace("\n", "").Replace(".", ",");
                            CNT.Write($"SENS:ACQ:APER 100ms");
                            str1 = CNT.Query("READ:FREQ? (@1)n").Replace("\n", "").Replace(".", ",");
                        }
                        else
                        {
                            str1 = CNT.Query("MEAS:FREQ? (@3)n").Replace("\n", "").Replace(".", ",");
                            CNT.Write($"SENS:ACQ:APER 100ms");
                            str1 = CNT.Query("READ:FREQ? (@3)n").Replace("\n", "").Replace(".", ",");
                        }
                    
                        str1 = (Convert.ToDouble(str1) * Math.Pow(10, -9)).ToString();

                        if (Math.Abs(((Form1.freqList_freq[i] - Convert.ToDouble(str1)) / (Form1.freqList_freq[i]))) <= 8 * Math.Pow(10, -6))
                        {
                            result = "Да";
                        }

                        else
                        {
                            result = "Нет";
                        }

                        sw.WriteLine(String.Format("{0,20}", (Form1.freqList_freq[i]).ToString()) + "\t" + String.Format("{0,20}", str1) + "\t" + String.Format(CultureInfo.InvariantCulture, "{0,20:0.###E+00}", (Form1.freqList_freq[i] - Convert.ToDouble(str1)) / (Form1.freqList_freq[i])) + "\t" + String.Format("{0,20}", "8E-06") + "\t" + String.Format("{0,20}", result));


                        m_DrawGraphicOnForm.DrawGraphic(false);

                    }

                }

                Form1.IsRun = false;
                sw.Close();
                System.Media.SystemSounds.Beep.Play(); System.Media.SystemSounds.Beep.Play(); System.Media.SystemSounds.Beep.Play();
                dr = MessageBox.Show("Процедура поверки по частоте закончена. Просмотреть результаты?", "Внимание", MessageBoxButtons.OKCancel);
                if (dr == DialogResult.OK)
                {
                    Process.Start(fileName);
                }

                //Освобождаем ресурсы
                m_DrawGraphicOnForm.DrawGraphic(true);
                m_DrawGraphicOnForm.EndCalculate();

                sw.Close();
            }

            catch (Exception ex)
            {
                //Освобождаем ресурсы
                m_DrawGraphicOnForm.DrawGraphic(true);
                m_DrawGraphicOnForm.EndCalculate();

                Form1.IsRun = false;
                System.Media.SystemSounds.Beep.Play(); System.Media.SystemSounds.Beep.Play(); System.Media.SystemSounds.Beep.Play();
                MessageBox.Show("Нарушение инструкции измерений, повторите попытку поверки по частоте еще раз", "Внимание");
            }

        }


        private void Run1()
        {
            try
            {
                DialogResult dr;

                try
                {
                    // Create driver instance
                    VAC = (MessageBasedSession)ResourceManager.GetLocalManager().Open(Form1.addr_zva);
                }
                catch
                {
                    MessageBox.Show("Проверьте физическое и программное подключение анализатора цепей к компьютеру", "Внимание", MessageBoxButtons.OK);
                    m_DrawGraphicOnForm.EndCalculate();
                    return;
                }

                VAC.Timeout = 20000;
                VAC.Write("*RST");

                VAC.Write("SWE:POIN 1000");

                VAC.Write("BWID 10");

                double[] freqList = { 0.1, 0.7, 2, 13, 24.0 };

                int[] dopusk = { -80, -110, -115, -110 };

                double[] result_noise = new double[6];

                string fileName = Form1.path + "\\" + String.Format("noise_{0}.txt", Form1.date.ToShortTimeString().Replace(":", " "));

                StreamWriter sw = new StreamWriter(FileStream.Null);

                try
                {
                    sw = new StreamWriter(fileName);
                }

                catch
                {
                    System.Media.SystemSounds.Beep.Play(); System.Media.SystemSounds.Beep.Play(); System.Media.SystemSounds.Beep.Play();
                    MessageBox.Show("Отсутствуют права на запись файла по указанному пути, обратитесь к администратору", "Внимание", MessageBoxButtons.OK);
                    if (VAC != null)
                    {
                        // Close the driver
                        VAC.Dispose();
                    }
                    m_DrawGraphicOnForm.EndCalculate();
                    return;
                }

                sw.WriteLine("Определение уровня собственных шумов анализатора");
                sw.WriteLine();
                sw.WriteLine(String.Format("{0,20}", "f1, ГГц") + "\t" + String.Format("{0,20}", "f2, ГГц, дБм") + "\t" + String.Format("{0,20}", "Шум, дБм") + "\t" + String.Format("{0,20}", "Допуск. (макс), дБм") + "\t" + String.Format("{0,20}", "Соответ.?"));

                sw.Flush();

                double[] data = null;
                double[] data_X = null;

                dr = MessageBox.Show(String.Format("Подключите согласованные нагрузки ко всем портам анализатора цепей, после этого нажмите ОК"), "Внимание", MessageBoxButtons.OKCancel);

                if (dr == DialogResult.Cancel)
                {
                    if (VAC != null)
                    {
                        // Close the driver
                        VAC.Dispose();
                    }

                    sw.Close();
                    m_DrawGraphicOnForm.EndCalculate();
                    return;
                }

                Form1.IsRun = true;

                for (int ii = 0; ii < Form1.count_port_noise; ii++)
                {
                    if (!Form1.IsRun)
                    {
                        sw.Close();
                        Form1.IsRun = false;

                        if (VAC != null)
                        {
                            // Close the driver
                            VAC.Dispose();
                        }

                        dr = MessageBox.Show("Процедура поверки по частоте закончена. Просмотреть результаты?", "Внимание", MessageBoxButtons.OKCancel);

                        if (dr == DialogResult.OK)
                        {
                            Process.Start(fileName);
                        }
                        m_DrawGraphicOnForm.EndCalculate();
                        return;
                    }

                    sw.WriteLine();
                    sw.WriteLine("Порт {0}", Form1.number_port_noise[ii]);
                    System.Media.SystemSounds.Beep.Play(); System.Media.SystemSounds.Beep.Play(); System.Media.SystemSounds.Beep.Play();

                    switch (Form1.number_port_noise[ii])
                    {
                        case (1):
                            VAC.Write(@"CALC1:PAR:MEAS 'Trc1', 'S21'");
                            break;

                        case (2):
                            VAC.Write(@"CALC1:PAR:MEAS 'Trc1', 'S12'");
                            break;

                        case (3):
                            VAC.Write(@"CALC1:PAR:MEAS 'Trc1', 'S43'");
                            break;

                        case (4):
                            VAC.Write(@"CALC1:PAR:MEAS 'Trc1', 'S34'");
                            break;
                    }

                    Thread.Sleep((int)(1.1 * Convert.ToDouble(VAC.Query("SWE:TIME?").Replace("\n", ""),provider) * 1000));

                    string[] str_data = VAC.Query("CALC1:DATA? FDAT").Split(',');

                    data = new double[str_data.Length];
                    data_X = new double[str_data.Length];

                    for (int iii = 0; iii < str_data.Length; iii++)
                    {
                        data[iii] = Convert.ToDouble(str_data[iii], provider);
                        data_X[iii] = freqList[0] + iii * (freqList[freqList.Length - 1] - freqList[0]) / 1000.0;
                    }

                    int number = 0;
                    double mean = 0;
                    int cout_mean = 0;

                    string results = "Нет";

                    for (int ll = 1; ll < freqList.Length; ll++)
                    {
                        mean = 0;
                        cout_mean = 0;

                        for (int l = number + 1; l < data_X.Length; l++)
                        {
                            if (data_X[l] >= freqList[ll - 1])
                            {
                                for (int j = number; j <= l; j++)
                                {
                                    mean += data[j];

                                    cout_mean++;
                                }

                                mean = mean / cout_mean;

                                mean -= 10;

                                number = l;

                                if (mean <= dopusk[ll - 1])
                                {
                                    results = "Да";
                                }
                                else
                                {
                                    results = "Нет";
                                }

                                break;
                            }
                        }

                        sw.WriteLine(String.Format("{0,20}", freqList[ll - 1].ToString()) + "\t" + String.Format("{0,20}", freqList[ll].ToString()) + "\t" + String.Format("{0,20}", (mean).ToString()) + "\t" + String.Format("{0,20}", (dopusk[ll - 1]).ToString()) + "\t" + String.Format("{0,20}", results));
                    }

                    sw.Flush();
                    m_DrawGraphicOnForm.DrawGraphic(false);

                }

                Form1.IsRun = false;
                sw.Close();
              

                //Освобождаем ресурсы
                m_DrawGraphicOnForm.DrawGraphic(true);
                m_DrawGraphicOnForm.EndCalculate();

                System.Media.SystemSounds.Beep.Play(); System.Media.SystemSounds.Beep.Play(); System.Media.SystemSounds.Beep.Play();
                dr = MessageBox.Show("Процедура поверки по уровню шумов закончена. Просмотреть результаты?", "Внимание", MessageBoxButtons.OKCancel);

                if (dr == DialogResult.OK)
                {
                    Process.Start(fileName);
                }
            }

            catch (Exception ex)
            {
                //Освобождаем ресурсы
                m_DrawGraphicOnForm.DrawGraphic(true);
                m_DrawGraphicOnForm.EndCalculate();

                Form1.IsRun = false;
                System.Media.SystemSounds.Beep.Play(); System.Media.SystemSounds.Beep.Play(); System.Media.SystemSounds.Beep.Play();
                MessageBox.Show("Нарушение инструкции измерений, повторите попытку поверки по уровню шумов еще раз", "Внимание");
            }
        }
    }
}
