using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class BulkLink : Form
    {
        ViewDetails v;
        string id;
        public BulkLink(ViewDetails refer, string ID)
        {
            InitializeComponent();
            v = refer;
            id = ID;
            for (int i = 0; i < v.m.table.Rows.Count; i++)
            {
                if (v.m.table.Rows[i].RowState != DataRowState.Deleted)
                {
                    comboBox2.Items.Add(v.m.table.Rows[i]["Asset Number"]);
                    comboBox1.Items.Add(v.m.table.Rows[i]["Asset Number"]);
                }
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
            int id1 = Main.assetNoToID(comboBox1.Text);
            int id2 = Main.assetNoToID(comboBox2.Text);
            if (id1 < 0 || id1 < 0)
                MessageBox.Show("Failed to Add");
            else
            {
                v.table2.Rows.Add(new Object[] { id1, id2 });
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.InitialDirectory = "./";
            openFileDialog1.Filter = "csv (*.csv)|*.csv|All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.RestoreDirectory = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                var csvRows = File.ReadAllLines(openFileDialog1.FileName);
                foreach (var l in csvRows)
                {
                    int id1 = Main.assetNoToID(l.Split(',')[0]);
                    int id2 = Main.assetNoToID(l.Split(',')[1]);
                    if (id1 < 0 || id1 < 0)
                        MessageBox.Show("Failed to Add: " + (id1 < 0 ? id1:id2) );
                    else
                    {
                        v.table2.Rows.Add(new Object[] { id1, id2 });
                    }
                }
            }
        }
    }
}
