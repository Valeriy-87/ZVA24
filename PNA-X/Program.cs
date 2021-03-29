using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PNA_X
{
    static class Program
    {
        /// <summary>
        /// Главная точка входа для приложения.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            bool onlyInstance;
            var mtx = new System.Threading.Mutex(true, "AppName", out onlyInstance);

            //проверяем что вторая копия программы не запущена
            if (onlyInstance)
            {
                Application.Run(new Form1());
            }
            else
            {
                MessageBox.Show("Приложение уже запущено", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
        }
    }
}
