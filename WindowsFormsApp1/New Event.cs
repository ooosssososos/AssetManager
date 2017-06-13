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
    public partial class New_Event : Form
    {
        ViewDetails v;
        string id;
        public New_Event(ViewDetails V, string ID)
        {
            v = V;
            id = ID;
            InitializeComponent();
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            v.table.Rows.Add(new Object[] {id, dateTimePicker1.Value.ToString("yyyy-MM-dd"), richTextBox1.Text});
            v.button3_Click(null,null);
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
