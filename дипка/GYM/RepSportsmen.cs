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
    public partial class RepSportsmen : MetroFramework.Forms.MetroForm
    {
        public RepSportsmen()
        {
            InitializeComponent();
        }

        private void RepSportsmen_Load(object sender, EventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "DsRepSportsmen.Спортсмен". При необходимости она может быть перемещена или удалена.

            this.reportViewer1.RefreshReport();
        }

        private void metroTile1_Click(object sender, EventArgs e)
        {
            this.Bounds = Screen.PrimaryScreen.Bounds;
            this.TopMost = true;
        }
    }
}
