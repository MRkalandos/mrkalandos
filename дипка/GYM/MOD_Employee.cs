using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MetroFramework.Forms;
using System.Data.OleDb;
using System.IO;

namespace GYM
{
    public partial class MOD_Employee : MetroForm
    {
        public MOD_Employee()
        {
            InitializeComponent();
        }

        private void MOD_Employee_Load(object sender, EventArgs e)
        {

        }

        private void metroTile1_Click(object sender, EventArgs e)
        {
            HeadForm HF = new HeadForm();
            if ((textBox1.Text == "") ||
                (textBox2.Text == "") ||
                (textBox3.Text == "") ||
                
                (maskedTextBox1.Text == "") ||
                (metroTextBox4.Text == "") ||
                (metroTextBox5.Text == "") ||
                (metroTextBox6.Text == "") ||
                (metroComboBox1.Text == ""))
            {
                MetroFramework.MetroMessageBox.Show(this, "\nНе все поля заполнены", "Корректность", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                OleDbConnection con1 = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=ISgym.mdb");
                con1.Open(); OleDbCommand sss1 = new OleDbCommand(@"select *  
                                                                      from [сотрудник] 
                                                                      where пароль=@st1 ", con1);
                sss1.Parameters.AddWithValue("st1", metroTextBox5.Text);
                sss1.ExecuteNonQuery();
                if (sss1.ExecuteScalar() != null)
                {
                    con1.Close();
                    MetroFramework.MetroMessageBox.Show(this, "\nТакой пароль уже существует", "Корректность", MessageBoxButtons.OK, MessageBoxIcon.Error); return;
                }
                else
                {
                    this.DialogResult = DialogResult.OK;
                }
            }
        }

        private void metroButton1_Click(object sender, EventArgs e)
        {
            OpenFileDialog filedialog = new OpenFileDialog();
            filedialog.Filter = "JPG (*.jpg)|*.jpg|PNG (*.png)|*.png";
            if (filedialog.ShowDialog() == DialogResult.OK)
            {
                string fileName = filedialog.SafeFileName;
                string sourcePath = filedialog.FileName;
                string targetPath = @"PhotoEmployee";
                if (!Directory.Exists(targetPath)) //Если папки нет...
                    Directory.CreateDirectory(targetPath); //...создадим ее
                string destFile = Path.Combine(targetPath, fileName);
                try
                {
                       File.Copy(sourcePath, destFile, true);
                        metroTextBox6.Text = Path.GetFileName(fileName);// то добавляем
                    }

                catch
                {
                    MetroFramework.MetroMessageBox.Show(this, "\nПапка задействована, фото может быть не установлено", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                     metroTextBox6.Text = Path.GetFileName(fileName);// то добавляем
                }
            } }

        private void metroTile2_Click(object sender, EventArgs e)
        {
            if (DialogResult.Yes == MetroFramework.MetroMessageBox.Show(this, "\nВы уверены что хотите выйти без сохранения", "Выход", MessageBoxButtons.YesNo, MessageBoxIcon.Warning))
            {
                Close();
            }
        }

        private void metroTextBox1_KeyPress(object sender, KeyPressEventArgs e)
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

        private void metroTextBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= 'А' && blockCifr <= 'я'))
            {
                if (e.KeyChar != (char)Keys.Back)
                {
                    e.Handled = true;
                    DialogResult result = MetroFramework.MetroMessageBox.Show(this, "\nНеверный тип данных", "Корректность", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void metroTextBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= 'А' && blockCifr <= 'я'))
            {
                if (e.KeyChar != (char)Keys.Back)
                {
                    e.Handled = true;
                    DialogResult result = MetroFramework.MetroMessageBox.Show(this, "\nНеверный тип данных", "Корректность", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void maskedTextBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= '0' && blockCifr <= '9'))
            {
                if (e.KeyChar != (char)Keys.Back)
                {
                    e.Handled = true;
                    DialogResult result = MetroFramework.MetroMessageBox.Show(this, "\nНеверный тип данных", "Корректность", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void metroTextBox5_KeyPress(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= '0' && blockCifr <= '9'))
            {
                if (e.KeyChar != (char)Keys.Back)
                {
                    e.Handled = true;
                    DialogResult result = MetroFramework.MetroMessageBox.Show(this, "\nНеверный тип данных", "Корректность", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void metroTextBox1_Click(object sender, EventArgs e)
        {
            if (((TextBox)sender).Text.Length == 1)
                ((TextBox)sender).Text = ((TextBox)sender).Text.ToUpper();
            ((TextBox)sender).Select(((TextBox)sender).Text.Length, 0);
        }

        private void metroTextBox2_Click(object sender, EventArgs e)
        {
            if (((TextBox)sender).Text.Length == 1)
                ((TextBox)sender).Text = ((TextBox)sender).Text.ToUpper();
            ((TextBox)sender).Select(((TextBox)sender).Text.Length, 0);
        }

        private void metroTextBox3_Click(object sender, EventArgs e)
        {
            if (((TextBox)sender).Text.Length == 1)
                ((TextBox)sender).Text = ((TextBox)sender).Text.ToUpper();
            ((TextBox)sender).Select(((TextBox)sender).Text.Length, 0);
        }

        private void metroButton2_Click(object sender, EventArgs e)
        {
            Money money = new Money();
            money.ShowDialog();
        }
    }
}
