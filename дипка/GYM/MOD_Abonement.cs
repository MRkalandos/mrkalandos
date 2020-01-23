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
    public partial class MOD_Abonement : MetroFramework.Forms.MetroForm
    {
        public MOD_Abonement()
        {
            InitializeComponent();
        }

        private void MOD_Abonement_Load(object sender, EventArgs e)
        {

        }

        private void metroTile1_Click(object sender, EventArgs e)
        {
            HeadForm HF = new HeadForm();
            if ((textBox1.Text == "") ||
                (textBox2.Text == "") ||
                (textBox3.Text == "") ||
                (metroComboBox1.Text == ""))
            {
                MetroFramework.MetroMessageBox.Show(this, "\nНе все поля заполнены", "Корректность", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                OleDbConnection con1 = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName + "/ISgym.mdb;Jet OLEDB:Database Password=316206");
                con1.Open(); OleDbCommand sss1 = new OleDbCommand(@"select *  
                                                                      from [Абонемент] 
                                                                      where Название=@st1 ", con1);
                sss1.Parameters.AddWithValue("st1", textBox1.Text);
                sss1.ExecuteNonQuery();
                if (sss1.ExecuteScalar() != null)
                {
                    con1.Close();
                    MetroFramework.MetroMessageBox.Show(this, "\nТакое название уже существует", "Корректность", MessageBoxButtons.OK, MessageBoxIcon.Error); return;
                }
                else
                {
                    this.DialogResult = DialogResult.OK;
                }
            }
        }

        private void metroTile2_Click(object sender, EventArgs e)
        {
            if (DialogResult.Yes == MetroFramework.MetroMessageBox.Show(this, "\nВы уверены что хотите выйти без сохранения", "Выход", MessageBoxButtons.YesNo, MessageBoxIcon.Warning))
            {
                Close();
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

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= '0' && blockCifr <= '9'))
            {
                if (e.KeyChar != (char)Keys.Back)
                {
                    e.Handled = true;
                    DialogResult result = MessageBox.Show("Неверный тип данных", "Корректность ввода", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= '0' && blockCifr <= '9'))
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
        void NewForm()
        {
            HeadForm form = new HeadForm();
            form.pictureBox2.Visible = true;

            form.metroTabControl1.SelectedTab = form.tabPage3;
        
            ((Control)form.EMPLtabPage6).Enabled = false;
            ((Control)form.tabPage5).Enabled = false;
            ((Control)form.tabPage4).Enabled = false;
            ((Control)form.tabPage29).Enabled = false;
            ((Control)form.tabPage13).Enabled = false;
            ((Control)form.tabPage30).Enabled = false;
            ((Control)form.tabPage27).Enabled = false;
            ((Control)form.tabPage26).Enabled = false;
            ((Control)form.tabPage25).Enabled = false;
            ((Control)form.tabPage24).Enabled = false;
            ((Control)form.tabPage22).Enabled = false;
            form.ShowDialog();
        }
      
        private void metroButton2_Click(object sender, EventArgs e)
        {
           new System.Threading.Thread(NewForm).Start();
        }
    }
}
