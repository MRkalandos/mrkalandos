using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GYM
{
   public static class ConnectionDb
    {
        public static string conString()
        {
            return (@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +
                    Directory.GetParent(Directory.GetCurrentDirectory()).Parent?.FullName +
                    "/ISgym.mdb;Jet OLEDB:Database Password=316206");
        }
    }
    }

