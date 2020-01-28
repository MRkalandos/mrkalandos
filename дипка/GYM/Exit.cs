using System;
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
            metroTile5.Focus();
        }

        private void metroTile6_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void metroTile5_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            metroLabel1.Text = DateTime.Now.ToString();
        }
    }
}
