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
    public partial class ReplaceAsset : Form
    {
        ViewDetails v;
        private Button button1;
        private ComboBox comboBox1;
        string id;
        public ReplaceAsset(ViewDetails refer, string ID)
        {
            InitializeComponent();
            v = refer;
            id = ID;
            for(int i = 0; i < v.m.table.Rows.Count; i++)
            {
                comboBox1.Items.Add(v.m.table.Rows[i]["Asset Number"]);
            }
            comboBox1.Sorted = true;
        }

        private void LinkAsset_Load(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void InitializeComponent()
        {
            this.button1 = new System.Windows.Forms.Button();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(12, 49);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(122, 23);
            this.button1.TabIndex = 0;
            this.button1.Text = "Replace";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click_1);
            // 
            // comboBox1
            // 
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Location = new System.Drawing.Point(13, 13);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(121, 21);
            this.comboBox1.TabIndex = 1;
            // 
            // ReplaceAsset
            // 
            this.ClientSize = new System.Drawing.Size(149, 91);
            this.Controls.Add(this.comboBox1);
            this.Controls.Add(this.button1);
            this.Name = "ReplaceAsset";
            this.Text = "Replace Asset";
            this.Load += new System.EventHandler(this.ReplaceAsset_Load);
            this.ResumeLayout(false);

        }

        private void ReplaceAsset_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            for (int i = 0; i < v.table2.Rows.Count; i++)
            {
                bool flag = false;
                if (v.table2.Rows[i].RowState != DataRowState.Deleted)
                {
                    //Console.Beep();
                    if (id == v.table2.Rows[i][0].ToString())
                    {
                        v.table2.Rows[i][0] = Main.assetNoToID(comboBox1.Text);
                        flag = true;
                    }
                    else if(id == v.table2.Rows[i][1].ToString())
                    {
                        v.table2.Rows[i][1] = Main.assetNoToID(comboBox1.Text);
                        flag = true;
                    }
                    if (flag)
                    {
                        //whats flag for???
                    }
                }

            }

            v.table.Rows.Add(new Object[] { Main.assetNoToID(comboBox1.Text), DateTime.Now.ToString("yyyy-MM-dd"), "Replaced " + Main.idToAssetNo(Int32.Parse(id)) });
            v.table.Rows.Add(new Object[] { id, DateTime.Now.ToString("yyyy-MM-dd"), "Replaced By " + comboBox1.Text });

            v.button3_Click(null, null);
            v.button4_Click(null, null);

            v.refresh2();
            v.Refresh();
            // new ViewDetails(v.m, id).Show();
            // v.Close();
            this.Close();
        }
    }
}
