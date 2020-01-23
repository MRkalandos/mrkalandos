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
    public partial class RepAbone_ent : MetroFramework.Forms.MetroForm
    {
        public RepAbone_ent()
        {
            InitializeComponent();
        }
        public string conString = (@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName + "/ISgym.mdb;Jet OLEDB:Database Password=316206");

        private void RepAbone_ent_Load(object sender, EventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "REPAbonement.Абонемент". При необходимости она может быть перемещена или удалена.
          
                   OleDbConnection conn = new OleDbConnection(conString);
            conn.Open();
            DataSet ds = new DataSet();

            OleDbDataAdapter adapter = new OleDbDataAdapter(String.Format(@"select * from абонемент"), conn);
            adapter.Fill(REPAbonement.Абонемент);
            conn.Close();
            this.reportViewer1.RefreshReport();
        }
    }
}
