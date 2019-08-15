using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SISOLVERTEST
{
    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            //Application.Run(new SISOLVERTEST());
            if(args.Length == 0)
            {
                Application.Run(new SISOLVERTEST());
                Application.Exit();
                //MessageBox.Show("外部登录");
            }
            else if (args.Length == 1)
            {
                Application.Run(new SISOLVERTEST(args[0]));
                Application.Exit();
            }
            else if (args.Length == 2)
            {
                Application.Run(new SISOLVERTEST(args[1]));
                Application.Exit();
            }
            //else if (args.Length == 0)
            //{
            //    Application.Run(new SISOLVERTEST("ME10N20JG6A3"));
            //}
            else 
            {
                Application.Exit();
                //MessageBox.Show("外部登录");
            }
        }
    }
}
