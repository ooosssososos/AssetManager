using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class newGroup : Form
    {
        public int ID { get; set; }
        Main refer;
        public newGroup(Main m)
        {
            InitializeComponent();
            refer = m;
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if(refer.conn.State == ConnectionState.Closed)
            refer.conn.Open();
            OleDbCommand c = new OleDbCommand("INSERT INTO [Groups] ([Group Name], [Date Created], [Description]) Values (@a,@b,@c)", refer.conn);
            c.Parameters.AddWithValue("@a", textBox1.Text);
            c.Parameters.AddWithValue("@b", dateTimePicker1.Text);
            c.Parameters.AddWithValue("@c", richTextBox1.Text);
            c.ExecuteNonQuery();
            c.CommandText = "Select @@Identity";
            ID = (int)c.ExecuteScalar();
            this.DialogResult = DialogResult.OK;
            Console.Out.WriteLine(ID);
            this.Close();
        }

        private void newGroup_Load(object sender, EventArgs e)
        {

        }
    }
}
