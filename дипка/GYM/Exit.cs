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
    public partial class Exit : MetroFramework.Forms.MetroForm
    {
        public Exit()
        {
            InitializeComponent();
        }

        private void Exit_Load(object sender, EventArgs e)
        {

        }

        private void metroTile6_Click(object sender, EventArgs e)
        {
            Close();

        }

        private void metroTile5_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
           
        }
    }
}
