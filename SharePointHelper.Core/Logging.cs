using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharePointHelper.Core
{
    class Logging
    {
        public static void LogMessage(string Message)
        {
            try
            {
                string line;
                string exePath = AppDomain.CurrentDomain.BaseDirectory;
                //string exePath = @"C:\Users\Administrator\Documents\Visual Studio 2013\Projects\SharePointHelper\SharePointHelper.Service\bin\Debug";

                line = string.Format("{0} {1}: {2}", DateTime.Now.ToLongTimeString(), DateTime.Now.ToLongDateString(), Message);

                using (StreamWriter file = new StreamWriter(@exePath + "\\" + "Log.txt", true))
                {
                    file.WriteLine(line);
                }
            }
            catch
            {

            }
        }
    }
}
