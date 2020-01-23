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
    public partial class MOD_SALE : MetroFramework.Forms.MetroForm
    {
        public MOD_SALE()
        {
            InitializeComponent();
        }
       
        private void VISITS_Load(object sender, EventArgs e)
        {

        }

        private void metroTile1_Click(object sender, EventArgs e)
        {
            if ((metroComboBox1.Text == "") ||
             (metroComboBox2.Text == "") ||
             (metroComboBox3.Text == ""))
            {
                MetroFramework.MetroMessageBox.Show(this, "\nНе все поля заполнены", "Корректность", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                this.DialogResult = DialogResult.OK;
                }
            }

        private void metroTile2_Click(object sender, EventArgs e)
        {
            if (DialogResult.Yes == MetroFramework.MetroMessageBox.Show(this, "\nВы уверены что хотите выйти без сохранения", "Выход", MessageBoxButtons.YesNo, MessageBoxIcon.Warning))
            {
                Close();
            }
        }
    }
    }

