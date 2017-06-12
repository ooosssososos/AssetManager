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
    public partial class Main : Form
    {
        public OleDbConnection conn;
        OleDbCommand cmd;
        //parameter from mdsaputra.udl
        private OleDbCommandBuilder oleCommandBuilder = null;
        OleDbDataAdapter da;
        public DataTable table = new DataTable();

        private String connParam = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\shihao\Desktop\Database11.accdb;Persist Security Info=False;";

        static Main m;

        public static string idToAssetNo(int id)
        {
            for(int i = 0; i < m.table.Rows.Count; i++)
            {
                if(!(m.table.Rows[i].RowState == DataRowState.Deleted))
                if(Int32.Parse(m.table.Rows[i]["ID"].ToString()) == id)
                {
                    return m.table.Rows[i]["Asset Number"].ToString();
                }
            }
            return "";
        }

        public DataRow findById(int id)
        {

                for (int i = 0; i < m.table.Rows.Count; i++)
                {
                    if (!(m.table.Rows[i].RowState == DataRowState.Deleted))
                        if (Int32.Parse(m.table.Rows[i]["ID"].ToString()) == id)
                        {
                            return m.table.Rows[i];
                        }
                }
                return null;

        }

        public static int assetNoToID(string id)
        {
            for (int i = 0; i < m.table.Rows.Count; i++)
            {

                if (!(m.table.Rows[i].RowState == DataRowState.Deleted))
                    if (m.table.Rows[i]["Asset Number"].ToString() == id)
                {
                    return Int32.Parse(m.table.Rows[i]["ID"].ToString());
                }
            }
            return -1;
        }

        public Main(String path)
        {
            connParam = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source="+ path + ";Persist Security Info=False;";
            conn = new OleDbConnection(connParam);
            InitializeComponent();
            m = this;
        }
        BindingSource bindingSource;
        private void Form1_Load(object sender, EventArgs e)
        {

            string strSQL = "SELECT * FROM assets";  //rename Sheet$ to yours sheet name (Code$ you said)
            cmd = new OleDbCommand(strSQL, conn);
            da = new OleDbDataAdapter(cmd);
            oleCommandBuilder = new OleDbCommandBuilder(da);
            oleCommandBuilder.QuotePrefix = "[";
            oleCommandBuilder.QuoteSuffix = "]";
            bindingSource = new BindingSource { DataSource = table };

            dataGridView1.DataSource = null;
            table.Clear();
            da.Fill(table);
            dataGridView1.DataSource = bindingSource;
            //this.AutoSize = true;
            // this.AutoSizeMode = AutoSizeMode.GrowAndShrink;
        }
        private void button1_Click(object sender, EventArgs e)
        {
            new newAsset(this).Show();
        }

        
        public void button2_Click(object sender, EventArgs e)
        {
            dataGridView1.EndEdit();
            if(da != null)
            da.Update(table);
            table.AcceptChanges();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
        }

        private void bindingSource1_CurrentChanged(object sender, EventArgs e)
        {

        }
        string search = "True";
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            search = "[Asset Number] LIKE '" + textBox1.Text + "*'";
            genRowFilter();
            dataGridView1.Refresh();
        }

        private void dataGridView1_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            new ViewDetails(this, dataGridView1.Rows[e.RowIndex].Cells["ID"].Value.ToString()).Show();
        }
        string surp = "True";
        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked) surp = "NOT [Surplus]";
            else surp = "True";
            genRowFilter();
            dataGridView1.Refresh();
        }

        private void genRowFilter()
        {
            table.DefaultView.RowFilter = "(" + surp.ToString() + ") AND (" + search.ToString() + ")";
        }

        private void dataGridView1_UserDeletingRow(object sender, DataGridViewRowCancelEventArgs e)
        {
            
        }

        private void Main_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (table.GetChanges() != null && MessageBox.Show("Save Changes if any?", "Confirm", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                button2_Click(null, null);
            }
        }

        private void dataGridView1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == 46)
                e.Handled = MessageBox.Show("Do you want really to delete the selected rows", "Confirm", MessageBoxButtons.OKCancel) != DialogResult.OK;
        }
    }
}
