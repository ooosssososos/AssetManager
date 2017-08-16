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
            a["Status"] = comboBox3.Text;
            a["Type"] = comboBox1.Text;
            a["Date Added"] = dateTimePicker1.Value.ToString("yyyy-MM-dd");
            a["Model"] = textBox6.Text;
            if(textBox8.Text != "")
            a["Memory"] = Int32.Parse(textBox8.Text);
            if (textBox7.Text != "")
                a["Hard Drive"] = Int32.Parse(textBox7.Text);
            a["Building"] = comboBox2.Text;
            a["OS"] = comboBox4.Text;
            a["On Network"] = checkBox1.Checked;
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

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }
    }
}
