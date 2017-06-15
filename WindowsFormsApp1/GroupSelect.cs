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
    public partial class GroupSelect : Form
    {
        public int GID { get; set; }
        public GroupSelect()
        {
            InitializeComponent();
        }

        private void GroupSelect_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                GID = Int32.Parse(textBox1.Text);
                this.DialogResult = DialogResult.OK;
                this.Close();
            }catch(Exception err)
            {

            }
        }
    }
}
