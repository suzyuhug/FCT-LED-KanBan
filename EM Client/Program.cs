using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace EM_Client
{
    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            bool createNew;
            using (System.Threading.Mutex m = new System.Threading.Mutex(true, Application.ProductName, out createNew))
            {
                if (createNew)
                {
                    Application.EnableVisualStyles();
                    Application.SetCompatibleTextRenderingDefault(false);
                    Application.Run(new Form1());
                }
                else
                {
                    MessageBox.Show("程序只能启动一次，有任务正在继续");
                }
            }         
        }
    }
}
