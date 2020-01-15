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
    public partial class AddEmployee : Form
    {
        public AddEmployee()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            HeadForm HF = new HeadForm();
            if ((textBox1.Text == "") ||
                (textBox2.Text == "") ||
                (textBox3.Text == "") ||
                (textBox4.Text == "") ||
                (maskedTextBox1.Text == "") ||
                (textBox5.Text == "") ||
                (comboBox1.Text == ""))
            {
                MessageBox.Show("Не все поля заполнены!");
            }
            else
            {
                OleDbConnection con1 = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=ISgym.mdb");
                con1.Open(); OleDbCommand sss1 = new OleDbCommand(@"select *  
                                                                      from [сотрудник] 
                                                                      where пароль=@st1 ", con1);
                sss1.Parameters.AddWithValue("st1", textBox5.Text);
                sss1.ExecuteNonQuery();
                if (sss1.ExecuteScalar() != null)
                {
                    con1.Close();
                    MessageBox.Show("Такой пароль существует"); return;
                }
                else
                {
                    this.DialogResult = DialogResult.OK;
                }
            }
        }
        private void AddEmployee_Load(object sender, EventArgs e)
        {
            this.зарплата_сотрудникаTableAdapter.Fill(this.dS_Money.Зарплата_сотрудника);

            //  this.ForeColor = System.Drawing.Color.Green;
            //this.COLO = System.Drawing.Color.Green;
        }


        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog filedialog = new OpenFileDialog();
            filedialog.Filter = "JPG (*.jpg)|*.jpg|PNG (*.png)|*.png"; 
            if (filedialog.ShowDialog() == DialogResult.OK)
            {
                string fileName = filedialog.SafeFileName;
                string sourcePath = filedialog.FileName;
                string targetPath = @"model";
                if (!Directory.Exists(targetPath)) //Если папки нет...
                Directory.CreateDirectory(targetPath); //...создадим ее
                string destFile = Path.Combine(targetPath, fileName);
               // File.Copy(sourcePath, destFile, true);
                 textBox6.Text= Path.GetFileName(fileName);// то добавляем
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (DialogResult.Yes == MessageBox.Show("Выйти без сохранения изменений?", "Подтверждение выхода", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation))
            {
                Close();
            }
            }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= 'А' && blockCifr <= 'я' ))
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
            if (!(blockCifr >= 'А' && blockCifr <= 'я' ))
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
            if (!(blockCifr >= 'А' && blockCifr <= 'я' ))
            {
                if (e.KeyChar != (char)Keys.Back)
                {
                    e.Handled = true;
                    DialogResult result = MessageBox.Show("Неверный тип данных", "Корректность ввода", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void textBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= 'А' && blockCifr <= 'я' || blockCifr <= ' '))
            {
                if (e.KeyChar != (char)Keys.Back)
                {
                    e.Handled = true;
                    DialogResult result = MessageBox.Show("Неверный тип данных", "Корректность ввода", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void maskedTextBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= '0' && blockCifr <= '9' || blockCifr <= '+'))
            {
                if (e.KeyChar != (char)Keys.Back)
                {
                    e.Handled = true;
                    DialogResult result = MessageBox.Show("Неверный тип данных", "Корректность ввода", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void textBox5_KeyPress(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= '0' && blockCifr <= '9' ))
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

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            if (((TextBox)sender).Text.Length == 1)
            ((TextBox)sender).Text = ((TextBox)sender).Text.ToUpper();
            ((TextBox)sender).Select(((TextBox)sender).Text.Length, 0);
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            if (((TextBox)sender).Text.Length == 1)
            ((TextBox)sender).Text = ((TextBox)sender).Text.ToUpper();
            ((TextBox)sender).Select(((TextBox)sender).Text.Length, 0);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Money money = new Money();
            money.ShowDialog();
        }

        private void AddEmployee_Shown(object sender, EventArgs e)
        {
            HeadForm head = new HeadForm();
            head.upd();
        }

        private void comboBox1_Click(object sender, EventArgs e)
        {
            this.зарплата_сотрудникаTableAdapter.Fill(this.dS_Money.Зарплата_сотрудника);
            //    bindingSource1.
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
         //   this.зарплата_сотрудникаTableAdapter.Update(this.dS_Money.Зарплата_сотрудника);
         //   this.зарплата_сотрудникаTableAdapter.Fill(this.dS_Money.Зарплата_сотрудника);

        }

        private void AddEmployee_FormClosed(object sender, FormClosedEventArgs e)
        {

        //    e.CloseReason = false
        }

        private void AddEmployee_FormClosing(object sender, FormClosingEventArgs e)
        {
         //   e.CloseReason. = false;
        }
    }
}
