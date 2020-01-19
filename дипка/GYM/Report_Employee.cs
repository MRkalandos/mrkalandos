using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
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

        private void Report_Employee_Load(object sender, EventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "DSREPemployee.vie1". При необходимости она может быть перемещена или удалена.
            this.vie1TableAdapter.Fill(this.DSREPemployee.vie1);
            this.reportViewer1.RefreshReport();
        }
    }
}
