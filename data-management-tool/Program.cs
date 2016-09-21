using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.IO;

namespace 无忧车秘神器_xp
{
    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            try
            {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
            }
            catch (Exception e)
            {
                FileStream fw = new FileStream("error.log", FileMode.OpenOrCreate);
                StreamWriter sw = new StreamWriter(fw);
                sw.WriteLine("错误信息：");
                sw.WriteLine(e.Message);
                sw.WriteLine("错误位置：");
                sw.WriteLine(e.StackTrace);
                sw.WriteLine("异常发生方法");
                sw.WriteLine(e.TargetSite.ToString());
                sw.Flush();
                sw.Close();
            }
            finally
            {
                Application.Exit();
            }
        }
    }
}
