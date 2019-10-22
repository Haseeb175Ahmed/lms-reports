using log4net;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AttendanceReport
{
    class CreateLogFiles
    {

        protected static readonly ILog log = LogManager.GetLogger(typeof(CreateLogFiles));

        public static void ErrorLog(string sErrMsg)
        {
            string fileName = DateTime.Now.Date.ToShortDateString().Replace('/', '-');
            string path = @"C:\AttendenceLog\logs\";

            if (!Directory.Exists(path))
            {
                DirectoryInfo di = Directory.CreateDirectory(path);
            }
            path = path + fileName + ".txt";
            //Check if the file exists
            if (!File.Exists(path))
            {
                // Create the file and use streamWriter to write text to it.
                //If the file existence is not check, this will overwrite said file.
                //Use the using block so the file can close and vairable disposed correctly
                using (StreamWriter writer = File.CreateText(path))
                {
                    writer.WriteLine(DateTime.Now.ToString() + " " + sErrMsg);
                }
            }
            else
            {
                using (StreamWriter writer = new StreamWriter(path, true))
                {
                    writer.WriteLine(DateTime.Now.ToString() + " " + sErrMsg);
                }
            }

        }

    }
}
