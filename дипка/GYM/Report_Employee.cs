using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GYM
{
    public partial class Report_Employee : MetroFramework.Forms.MetroForm
    {
        public Report_Employee()
        {
            InitializeComponent();
        }
        public string conString = (@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName + "/ISgym.mdb;Jet OLEDB:Database Password=316206");

        private void Report_Employee_Load(object sender, EventArgs e)
        {
            try
            {
                OleDbConnection conn = new OleDbConnection(conString);
                conn.Open();
                DataSet ds = new DataSet();

                OleDbDataAdapter adapter = new OleDbDataAdapter(String.Format(@"select * from vie1"), conn);
                adapter.Fill(this.DSREPemployee.vie1);
                conn.Close();
                this.reportViewer1.RefreshReport();
            }
            catch {
            }
        }
    }
}
