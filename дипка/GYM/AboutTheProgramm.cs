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
    public partial class AboutTheProgramm : MetroFramework.Forms.MetroForm
    {
        public AboutTheProgramm()
        {
            InitializeComponent();
        }

        private void Progruzka_Load(object sender, EventArgs e)
        {
            

        }

        private void metroTile1_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
