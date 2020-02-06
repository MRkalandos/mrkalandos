using System;
using System.IO;
using System.Windows.Forms;
using MetroFramework;

namespace GYM
{
    public partial class ModSale : MetroFramework.Forms.MetroForm
    {
        private const string TitleException = "Ошибка";

        public ModSale()
        {
            InitializeComponent();
            this.KeyPreview = true;
        }

        private void VISITS_Load(object sender, EventArgs e)
        {
            FocusMe();
        }

        private void metroTile1_Click(object sender, EventArgs e)
        {
            if ((metroComboBox1.Text == "") ||
                (metroComboBox2.Text == "") ||
                (metroComboBox3.Text == ""))
            {
                MetroMessageBox.Show(this, "\nНе все поля заполнены", TitleException,
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                this.DialogResult = DialogResult.OK;
            }
        }

        private void metroTile2_Click(object sender, EventArgs e)
        {
            if (DialogResult.Yes == MetroFramework.MetroMessageBox.Show(this,
                    "\nВы уверены что хотите выйти без сохранения", "Выход", MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning))
            {
                Close();
            }
            else
            {
                FocusMe();
            }
        }

        private static void NewForm1()
        {
            var form = new HeadForm {pictureBox2 = {Visible = true}};

            form.metroTabControl1.SelectedTab = form.tabPage1;
            form.metroTabControl2.SelectedTab = form.EMPLtabPage6;
            ((Control) form.tabPage5).Enabled = false;
            ((Control) form.tabPage4).Enabled = false;
            ((Control) form.tabPage3).Enabled = false;
            ((Control) form.tabPage2).Enabled = false;
            ((Control) form.tabPage13).Enabled = false;
            ((Control) form.tabPage8).Enabled = false;
            ((Control) form.tabPage9).Enabled = false;
            ((Control) form.tabPage10).Enabled = false;
            ((Control) form.tabPage11).Enabled = false;
            ((Control) form.tabPage12).Enabled = false;
            form.ShowInTaskbar = false;
            form.ControlBox = false;
            form.ShowDialog();
        }

        private void metroButton2_Click(object sender, EventArgs e)
        {
            new System.Threading.Thread(NewForm1).Start();
        }

        private static void NewForm2()
        {
            var form = new HeadForm {pictureBox2 = {Visible = true}};
            form.metroTabControl1.SelectedTab = form.tabPage1;
            form.metroTabControl2.SelectedTab = form.tabPage4;
            ((Control) form.tabPage5).Enabled = false;
            ((Control) form.EMPLtabPage6).Enabled = false;
            ((Control) form.tabPage3).Enabled = false;
            ((Control) form.tabPage2).Enabled = false;
            ((Control) form.tabPage13).Enabled = false;
            ((Control) form.tabPage17).Enabled = false;
            ((Control) form.tabPage16).Enabled = false;
            ((Control) form.tabPage15).Enabled = false;
            ((Control) form.tabPage14).Enabled = false;
            form.ControlBox = false;
            form.ShowInTaskbar = false;
            form.ShowDialog();
        }

        private void metroButton1_Click(object sender, EventArgs e)
        {
            new System.Threading.Thread(NewForm2).Start();
        }

        private static void NewForm()
        {
            var form = new HeadForm {pictureBox2 = {Visible = true}};
            form.metroTabControl1.SelectedTab = form.tabPage2;
            form.metroTabControl2.SelectedTab = form.tabPage29;
            ((Control) form.tabPage1).Enabled = false;
            ((Control) form.tabPage3).Enabled = false;
            ((Control) form.tabPage13).Enabled = false;
            ((Control) form.tabPage30).Enabled = false;
            ((Control) form.tabPage34).Enabled = false;
            ((Control) form.tabPage33).Enabled = false;
            ((Control) form.tabPage32).Enabled = false;
            ((Control) form.tabPage31).Enabled = false;
            form.ShowInTaskbar = false;
            form.ControlBox = false;
            form.ShowDialog();
        }

        private void metroButton3_Click(object sender, EventArgs e)
        {
            new System.Threading.Thread(NewForm).Start();
        }

        private void MOD_SALE_Shown(object sender, EventArgs e)
        {
            FocusMe();
        }

        private void MOD_SALE_Activated(object sender, EventArgs e)
        {
            FocusMe();
        }

        private void MOD_SALE_KeyDown(object sender, KeyEventArgs e)
        {
            switch (e.KeyCode)
            {
                case Keys.F7:
                    metroButton1.PerformClick();
                    break;
                case Keys.F8:
                    metroButton3.PerformClick();
                    break;
                case Keys.F6:
                    metroButton2.PerformClick();
                    break;
                case Keys.F5:
                    metroTile1.PerformClick();
                    break;
                case Keys.F1:
                    pictureBox1_Click(this, e);
                    break;
                case Keys.Escape:
                    metroTile2.PerformClick();
                    break;
            }
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            try
            {
                if (File.Exists("Help/Help.chm"))
                {
                    Help.ShowHelp(null, "Help/Help.chm");
                }
                else
                {
                    MetroMessageBox.Show(this, "Файл не найден", TitleException, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
            catch (Exception exception)
            {
                MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                HelperLog.Write(exception.Message);
            }
            finally
            {
                FocusMe();
            }
        }

        private void ModSale_Click(object sender, EventArgs e)
        {
            FocusMe();
        }
    }
}

