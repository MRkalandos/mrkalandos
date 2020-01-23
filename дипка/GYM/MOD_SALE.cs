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


        void NewForm1()
        {
            HeadForm form = new HeadForm();
            form.pictureBox2.Visible = true;

            form.metroTabControl1.SelectedTab = form.tabPage1;
            form.metroTabControl2.SelectedTab = form.EMPLtabPage6;
            ((Control)form.tabPage5).Enabled = false;
            ((Control)form.tabPage4).Enabled = false;
            ((Control)form.tabPage3).Enabled = false;
            ((Control)form.tabPage2).Enabled = false;
            ((Control)form.tabPage13).Enabled = false;
            ((Control)form.tabPage8).Enabled = false;
            ((Control)form.tabPage9).Enabled = false;
            ((Control)form.tabPage10).Enabled = false;
            ((Control)form.tabPage11).Enabled = false;
            ((Control)form.tabPage12).Enabled = false;
            // new System.Threading.Thread(NewForm1).Start();
            form.ShowDialog();
        }

        private void metroButton2_Click(object sender, EventArgs e)
        {
            new System.Threading.Thread(NewForm1).Start();
        }


        void NewForm2()
        {
            HeadForm form = new HeadForm();
            form.pictureBox2.Visible = true;

            form.metroTabControl1.SelectedTab = form.tabPage1;
            form.metroTabControl2.SelectedTab = form.tabPage4;
            ((Control)form.tabPage5).Enabled = false;
            ((Control)form.EMPLtabPage6).Enabled = false;
            ((Control)form.tabPage3).Enabled = false;
            ((Control)form.tabPage2).Enabled = false;
            ((Control)form.tabPage13).Enabled = false;
            ((Control)form.tabPage17).Enabled = false;
            ((Control)form.tabPage16).Enabled = false;
            ((Control)form.tabPage15).Enabled = false;
            ((Control)form.tabPage14).Enabled = false;
         
            // new System.Threading.Thread(NewForm2).Start();
            form.ShowDialog();
        }
        private void metroButton1_Click(object sender, EventArgs e)
        {
            new System.Threading.Thread(NewForm2).Start();
        }
        void NewForm()
        {
            HeadForm form = new HeadForm();
            form.pictureBox2.Visible = true;

            form.metroTabControl1.SelectedTab = form.tabPage2;
            form.metroTabControl2.SelectedTab = form.tabPage29;


            ((Control)form.tabPage1).Enabled = false;
            ((Control)form.tabPage3).Enabled = false;
            ((Control)form.tabPage13).Enabled = false;
            ((Control)form.tabPage30).Enabled = false;
            //((Control)form.tabPage42).Enabled = false;
            ((Control)form.tabPage34).Enabled = false;
            ((Control)form.tabPage33).Enabled = false;
            ((Control)form.tabPage32).Enabled = false;
            ((Control)form.tabPage31).Enabled = false;
            form.ShowDialog();
        }
        private void metroButton3_Click(object sender, EventArgs e)
        {
            new System.Threading.Thread(NewForm).Start();
        }
    }
    }

