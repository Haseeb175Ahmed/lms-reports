using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AttendanceReport
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            SplashScreen splash = new SplashScreen();


            Task dbInitializerTask = new Task(() =>
            {
                try
                {
                    EFERTDbUtility.InitializeDatabases();
                    splash.Invoke(new Action(() => { splash.Close(); }));

                }
                catch (Exception)
                {

                    throw;
                }
            });

            dbInitializerTask.Start();

            splash.ShowDialog();
            Application.Run(new LoginForm());
        }
    }
}
