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
    public partial class New_GEvent : Form
    {
        GroupManager v;
        string Gid;
        public New_GEvent(GroupManager V, string GID)
        {
            v = V;
            Gid = GID;
            InitializeComponent();
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            v.table2.Rows.Add(new Object[] {Gid, dateTimePicker1.Value.ToString("yyyy-MM-dd"), richTextBox1.Text});
            v.button3_Click(null, null);
            this.Close();
        }   

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {

        }

        private void New_Event_Load(object sender, EventArgs e)
        {

        }
    }
}
