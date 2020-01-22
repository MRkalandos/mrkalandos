﻿using System;
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
    public partial class RepSportsmen : MetroFramework.Forms.MetroForm
    {
        public RepSportsmen()
        {
            InitializeComponent();
        }
        public string conString = (@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName + "/ISgym.mdb;Jet OLEDB:Database Password=316206");

        private void RepSportsmen_Load(object sender, EventArgs e)
        {
                  OleDbConnection conn = new OleDbConnection(conString);
            conn.Open();
            DataSet ds = new DataSet();

            OleDbDataAdapter adapter = new OleDbDataAdapter(String.Format(@"select * from vie2"), conn);
            adapter.Fill(DSREPSportsmenVIE2.vie2);
              conn.Close();
            this.reportViewer1.RefreshReport();
        }

        private void metroTile1_Click(object sender, EventArgs e)
        {
            
        }
    }
}
