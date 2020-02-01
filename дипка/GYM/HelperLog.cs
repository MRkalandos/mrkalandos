using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GYM
{
    public static class HelperLog
        {
            public static void Write(string exceptionMessage)
            {
                File.AppendAllText(Path(), $@"{GetDate()}:{exceptionMessage}");
            }
            private static string GetDate()
            {
                return DateTime.Now.ToString("dd MMMM yyyy | HH:mm:ss");
            }
            private static string Path()
            {
                return Directory.GetCurrentDirectory() + @"\" + "LOG/Log.txt";
            }
        }
}
