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
    public partial class Report : Form
    {
        public Report()
        {
            InitializeComponent();
        }

        private void Report_Load(object sender, EventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "DS_Money.Зарплата_сотрудника". При необходимости она может быть перемещена или удалена.
            this.Зарплата_сотрудникаTableAdapter.Fill(this.DS_Money.Зарплата_сотрудника);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "DSEmployee.Сотрудник". При необходимости она может быть перемещена или удалена.
            this.СотрудникTableAdapter.Fill(this.DSEmployee.Сотрудник);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "DSEmployee.Сотрудник". При необходимости она может быть перемещена или удалена.
            //   this.СотрудникTableAdapter.Fill(this.//DSEmployee.Сотрудник);

            // this.reportViewer1.RefreshReport();
            // this.reportViewer2.RefreshReport();
            this.reportViewer1.RefreshReport();
            //this.reportViewer2.RefreshReport();
        }

        private void reportViewer1_Load(object sender, EventArgs e)
        {

        }

        private void bindingSource1_CurrentChanged(object sender, EventArgs e)
        {

        }
    }
}
