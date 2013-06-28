using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApplication1test
{
    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            startForm sf = new startForm();
            if (sf.ShowDialog() == DialogResult.OK)
            {
                Application.Run(new mainForm(sf.getAddress(), sf.getNewSmry()));
            }
            //Application.Run(new startForm());            
        }
    }
}
