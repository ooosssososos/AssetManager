using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class newAsset : Form
    {
        Main m;
        public newAsset(Main refer)
        {
            m = refer;
            InitializeComponent();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            DataRow a = m.table.NewRow();
            a["Asset Number"] = textBox1.Text;
            a["User"] = textBox2.Text;
            a["Location"] = textBox3.Text;
            a["S/N"] = textBox4.Text;
            a["Manufacturer"] = textBox5.Text;
            a["Description"] = richTextBox1.Text;
            a["Surplus"] = checkBox1.Checked;
            a["Type"] = comboBox1.Text;
            a["Date Added"] = dateTimePicker1.Value.ToString("yyyy-MM-dd");
            a["Model"] = textBox6.Text;
            a["Memory (GB)"] = textBox8.Text;
            a["HDD (GB)"] = textBox7.Text;
            a["Building"] = comboBox2.Text;
            m.table.Rows.Add(a);
            m.button2_Click(null, null);
            m.refresh2();
            this.Close();

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {

        }

        private void label8_Click(object sender, EventArgs e)
        {

        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
