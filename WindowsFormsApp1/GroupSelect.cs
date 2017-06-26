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
    public partial class GroupSelect : Form
    {
        public int GID { get; set; }
        public GroupSelect()
        {
            InitializeComponent();
            if(Main.con.State == ConnectionState.Closed)
            Main.con.Open();
            string strSQL2 = "SELECT * FROM [Groups]";  //rename Sheet$ to yours sheet name (Code$ you said)
            OleDbCommand cmd2 = new OleDbCommand(strSQL2, Main.con);
            using (var Reader = cmd2.ExecuteReader())
            {
                while (Reader.Read())
                {
                    comboBox1.Items.Add(Reader.GetInt32(0) + ", " + Reader.GetString(1));
                }
            }
            
        }

        private void GroupSelect_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                GID = Int32.Parse(comboBox1.Text.Split(',')[0]);
                this.DialogResult = DialogResult.OK;
                this.Close();
            }catch(Exception err)
            {

            }
        }
    }
}
