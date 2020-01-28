using System;
using System.IO;
using System.Windows.Forms;

namespace GYM
{
    public partial class ModSale : MetroFramework.Forms.MetroForm
    {
        private const string TitleException = "Ошибка";
        private readonly string _dateLog = DateTime.Now.ToString("dd MMMM yyyy | HH:mm:ss");
        private readonly string _fileNameLog = Directory.GetCurrentDirectory() + @"\" + "LOG/Mod_sale.txt";

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
                MetroFramework.MetroMessageBox.Show(this, "\nНе все поля заполнены", TitleException,
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
                if (File.Exists("Help/Mod_sale.chm"))
                {
                    Help.ShowHelp(null, "Help/Mod_sale.chm");
                }
                else
                {
                    MetroFramework.MetroMessageBox.Show(this, "Файл не найден", TitleException,MessageBoxButtons.OK,MessageBoxIcon.Error);
                    FocusMe();
                }
            }
            catch (Exception exception)
            {
                MetroFramework.MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK, MessageBoxIcon.Error);
                if (File.Exists(_fileNameLog) != true)
                {
                    using (var sw =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                    {
                        sw.WriteLine(_dateLog);
                        sw.WriteLine(exception.Message);
                        FocusMe();
                    }
                }
                else
                {
                    using (var sw =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                    {
                        (sw.BaseStream).Seek(0, SeekOrigin.End);
                        sw.WriteLine(_dateLog);
                        sw.WriteLine(exception.Message);
                        FocusMe();
                    }
                }
            }
        }
    }
}

