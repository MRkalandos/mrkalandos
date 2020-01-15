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
    public partial class AddMoney : Form
    {
        public AddMoney()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (DialogResult.Yes == MessageBox.Show("Выйти без сохранения изменений?", "Подтверждение выхода", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation))
            {
                Close();
            }
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
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

        private void button1_Click(object sender, EventArgs e)
        {
            HeadForm HF = new HeadForm();
            if ((textBox1.Text == "") )
              
            {
                MessageBox.Show("Не все поля заполнены!");
            }
            else
            {
                this.DialogResult = DialogResult.OK;
            }
        }
    }
}
