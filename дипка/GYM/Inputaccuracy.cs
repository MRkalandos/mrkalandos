using System;
using System.Windows.Forms;
using MetroFramework;

namespace GYM
{
    public class Inputaccuracy
    {
        private const string Message = @"Неверный тип данных";
        private const string Title = @"Корректность ввода";

        public void ModTrenerInputaccuracyNumeral(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= '0' && blockCifr <= '9'))
            {
                if (e.KeyChar != (char) Keys.Back)
                {
                    e.Handled = true;
                    MetroMessageBox.Show(new ModTrener(), Message, Title,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
        }

        public void ModEmployeeInputaccuracyNumeral(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= '0' && blockCifr <= '9'))
            {
                if (e.KeyChar != (char) Keys.Back)
                {
                    e.Handled = true;
                    MetroMessageBox.Show(new ModEmployee(), Message, Title,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
        }

        public void HeadInputaccuracyNumeral(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= '0' && blockCifr <= '9'))
            {
                if (e.KeyChar != (char) Keys.Back)
                {
                    e.Handled = true;
                    MetroMessageBox.Show(new HeadForm(), Message, Title,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
        }

        public void ModAbonementInputaccuracyNumeral(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= '0' && blockCifr <= '9'))
            {
                if (e.KeyChar != (char) Keys.Back)
                {
                    e.Handled = true;
                    MetroMessageBox.Show(new ModAbonement(), Message, Title,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
        }

        public void HeadInputaccuracyLetter(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= 'А' && blockCifr <= 'я'))
            {
                if (e.KeyChar != (char) Keys.Back)
                {
                    e.Handled = true;
                    MetroMessageBox.Show(new HeadForm(), Message, Title,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
        }

        public void ModSportsmenInputaccuracyLetter(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= 'А' && blockCifr <= 'я'))
            {
                if (e.KeyChar != (char) Keys.Back)
                {
                    e.Handled = true;
                    MetroMessageBox.Show(new ModSportsmen(), Message, Title,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
        }

        public void ModEmployeeInputaccuracyLetter(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= 'А' && blockCifr <= 'я'))
            {
                if (e.KeyChar != (char) Keys.Back)
                {
                    e.Handled = true;
                    MetroMessageBox.Show(new ModEmployee(), Message, Title,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
        }


        public void UpperLetter(object sender, EventArgs e)
        {
            if (((TextBox) sender).Text.Length == 1)
                ((TextBox) sender).Text = ((TextBox) sender).Text.ToUpper();
            ((TextBox) sender).Select(((TextBox) sender).Text.Length, 0);
        }

        public void ModTrenerInputaccuracyLetter(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= 'А' && blockCifr <= 'я'))
            {
                if (e.KeyChar != (char) Keys.Back)
                {
                    e.Handled = true;
                    MetroMessageBox.Show(new ModTrener(), Message, Title, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
        }
    }
}
