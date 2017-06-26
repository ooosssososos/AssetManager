using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public static class PasswordPrompt
    {
        public static string[] ShowDialog()
        {
            Form prompt = new Form()
            {
                Width = 500,
                Height = 150,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                Text = "Login prompt",
                StartPosition = FormStartPosition.CenterScreen
            };
            TextBox textLabel = new TextBox() { Left = 50, Top = 20, Width = 400, };
            TextBox textBox = new TextBox() { Left = 50, Top = 50, Width = 400 };
            textBox.UseSystemPasswordChar = true;
            Button confirmation = new Button() { Text = "Ok", Left = 350, Width = 100, Top = 70, DialogResult = DialogResult.OK };
            confirmation.Click += (sender, e) => { prompt.Close(); };
            prompt.Controls.Add(textLabel);
            prompt.Controls.Add(textBox);
            prompt.Controls.Add(confirmation);
            prompt.AcceptButton = confirmation;
            if (prompt.ShowDialog() == DialogResult.OK)
            {
                string[] ret = { textLabel.Text, textBox.Text };
                return ret;

            }
            return null;
        }
    }
}
