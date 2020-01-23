using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GYM
{
    public partial class MOD_Trenerovka : MetroFramework.Forms.MetroForm
    {
        public MOD_Trenerovka()
        {
            InitializeComponent();
        }

        private void MOD_Trenerovka_Load(object sender, EventArgs e)
        {

        }

        private void metroTile2_Click(object sender, EventArgs e)
        {
            if (DialogResult.Yes == MetroFramework.MetroMessageBox.Show(this, "\nВы уверены что хотите выйти без сохранения", "Выход", MessageBoxButtons.YesNo, MessageBoxIcon.Warning))
            {
                Close();
            }
        }

        private void metroTile1_Click(object sender, EventArgs e)
        {
            HeadForm HF = new HeadForm();
            if ((textBox1.Text == "") ||
                (metroComboBox1.Text == "") ||
                (metroComboBox2.Text == ""))
               
            {
                MetroFramework.MetroMessageBox.Show(this, "\nНе все поля заполнены", "Корректность", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                this.DialogResult = DialogResult.OK;
            }
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= 'А' && blockCifr <= 'я'))
            {
                if (e.KeyChar != (char)Keys.Back)
                {
                    e.Handled = true;
                    DialogResult result = MessageBox.Show("Неверный тип данных", "Корректность ввода", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (((TextBox)sender).Text.Length == 1)
                ((TextBox)sender).Text = ((TextBox)sender).Text.ToUpper();
            ((TextBox)sender).Select(((TextBox)sender).Text.Length, 0);
        }

        private void metroButton1_Click(object sender, EventArgs e)
        {
            View_trenerovki view = new View_trenerovki();
            view.Show();
        }
        void NewForm()
        {
            HeadForm form = new HeadForm();
            form.pictureBox2.Visible = true;
            
            form.metroTabControl1.SelectedTab = form.tabPage1;
            form.metroTabControl2.SelectedTab = form.tabPage5;
            ((Control)form.tabPage2).Enabled = false;
            ((Control)form.tabPage3).Enabled = false;
            ((Control)form.tabPage13).Enabled = false;
            ((Control)form.EMPLtabPage6).Enabled = false;
            ((Control)form.tabPage4).Enabled = false;
            ((Control)form.tabPage20).Enabled = false;
            ((Control)form.tabPage21).Enabled = false;
            ((Control)form.tabPage18).Enabled = false;
            ((Control)form.tabPage6).Enabled = false;
            form.ShowDialog();
           

        }
        public void metroButton2_Click(object sender, EventArgs e)
        {
            new System.Threading.Thread(NewForm).Start();
        }
      





    }

}
    

