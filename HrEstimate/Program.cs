using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace HrEstimate
{
    static class Program
    {
        /// <summary>
        /// 應用程式的主要進入點。
        /// </summary>
        [STAThread]
        static void Main()
        {
            //處理未捕捉的例外
            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
            //處理UI執行緒錯誤
            Application.ThreadException += new System.Threading.ThreadExceptionEventHandler(Application_ThreadException);
            //非處理UI執行緒錯誤
            AppDomain.CurrentDomain.UnhandledException += new UnhandledExceptionEventHandler(CurrentDomain_UnhandledException);
            
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }

        //處理UI執行緒錯誤
        static void Application_ThreadException(object sender, System.Threading.ThreadExceptionEventArgs e)
        {
            string msg;
            Exception err = e.Exception as Exception;
            if (err != null)
            {
                if (err.Message.Contains("找不到或無法存取伺服器"))
                {
                    msg = string.Format("網路連線失敗，請離開目前作業重新操作\n");
                }
                else
                {
                    msg = string.Format("發生應用程式例外，請離開目前作業重新操作\n並聯絡資訊部車機處理人員\n");
                }
            }
            else
            {
                msg = string.Format("發生應用程式例外，請離開目前作業重新操作\n並聯絡資訊部車機處理人員\n");
            }
            MessageBox.Show(err.StackTrace+err.Source+err.Message);
        }

        //非處理UI執行緒錯誤
        static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            string msg;
            Exception err = e.ExceptionObject as Exception;
            if (err != null)
            {
                if (err.Message.Contains("找不到或無法存取伺服器"))
                {
                    msg = string.Format("網路連線失敗，請離開目前作業重新操作\n");
                }
                else
                {
                    msg = string.Format("發生應用程式例外，請離開目前作業重新操作\n並聯絡資訊部車機處理人員\n");
                }
            }
            else
            {
                msg = string.Format("Application UnhandleException，請離開目前作業重新操作\n並聯絡資訊部車機處理人員:{0}", e);
            }
            MessageBox.Show(err.StackTrace + err.Source + err.Message);
        }
    }
}
