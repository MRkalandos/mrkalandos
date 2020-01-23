using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GYM
{
    public partial class MOD_VISITS : MetroFramework.Forms.MetroForm
    {
        public MOD_VISITS()
        {
            InitializeComponent();
        }

        private void MOD_VISITS_Load(object sender, EventArgs e)
        {

        }

        private void metroTile1_Click(object sender, EventArgs e)
        {
            if ((metroDateTime1.Text == "") ||
                (metroComboBox1.Text == ""))
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

        void NewForm2()
        {
            HeadForm form = new HeadForm();
            form.pictureBox2.Visible = true;
            form.metroTabControl1.SelectedTab = form.tabPage2;
            form.metroTabControl2.SelectedTab = form.tabPage30;


            ((Control)form.tabPage1).Enabled = false;
            ((Control)form.tabPage3).Enabled = false;
            ((Control)form.tabPage13).Enabled = false;
            ((Control)form.tabPage29).Enabled = false;
            //((Control)form.tabPage42).Enabled = false;
            ((Control)form.tabPage40).Enabled = false;
            ((Control)form.tabPage39).Enabled = false;
            ((Control)form.tabPage38).Enabled = false;
            ((Control)form.tabPage37).Enabled = false;
            ((Control)form.tabPage36).Enabled = false;
            form.ShowDialog();
        }
        private void metroButton2_Click(object sender, EventArgs e)
        {
            new System.Threading.Thread(NewForm2).Start();
        }
    }
    
}
