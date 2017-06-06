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
    public partial class LinkAsset : Form
    {
        ViewDetails v;
        string id;
        public LinkAsset(ViewDetails refer, string ID)
        {
            InitializeComponent();
            v = refer;
            id = ID;
            for(int i = 0; i < v.m.table.Rows.Count; i++)
            {
                comboBox2.Items.Add(v.m.table.Rows[i]["Asset Number"]);
                comboBox1.Items.Add(v.m.table.Rows[i]["Asset Number"]);
            }
            comboBox1.Sorted = true;
            comboBox2.Sorted = true;
            comboBox1.SelectedIndex = comboBox1.FindString(Main.idToAssetNo(Int32.Parse(id)));
        }

        private void LinkAsset_Load(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            v.table2.Rows.Add(new Object[] { Main.assetNoToID(comboBox1.Text), Main.assetNoToID(comboBox2.Text)});
            v.button4_Click(null,null);
            this.Close();
        }
    }
}
