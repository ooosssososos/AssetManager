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
    public partial class ViewDetails : Form
    {
        private OleDbConnection conn;
        private OleDbConnection conn2;
        OleDbCommand cmd;
        OleDbCommand cmd2;
        //parameter from mdsaputra.udl
        private OleDbCommandBuilder oleCommandBuilder = null;
        OleDbDataAdapter da;
        public DataTable table = new DataTable();
        private OleDbCommandBuilder oleCommandBuilder2 = null;
        OleDbDataAdapter da2;
        public DataTable table2 = new DataTable();

        private String connParam = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\shihao\Desktop\Database11.accdb;Persist Security Info=False;";
        String id;
        public Main m;
        public ViewDetails(Main refer, String ID)
        {

            InitializeComponent();
            id = ID;
            m = refer;
            
                        conn = new OleDbConnection(connParam);
                        string strSQL = "SELECT * FROM History WHERE id = " + id + "";  //rename Sheet$ to yours sheet name (Code$ you said)
                        cmd = new OleDbCommand(strSQL, conn);
                        da = new OleDbDataAdapter(cmd);
                        oleCommandBuilder = new OleDbCommandBuilder(da);
                        oleCommandBuilder.QuotePrefix = "[";
                        oleCommandBuilder.QuoteSuffix = "]";
                        button3_Click(null, null);

            conn2 = new OleDbConnection(connParam);
            string strSQL2 = "SELECT * FROM relations WHERE asset1 = " + id + " OR asset2 = " + id+"";  //rename Sheet$ to yours sheet name (Code$ you said)
            cmd2 = new OleDbCommand(strSQL2, conn2);
            da2 = new OleDbDataAdapter(cmd2);
            oleCommandBuilder2 = new OleDbCommandBuilder(da2);
            oleCommandBuilder2.QuotePrefix = "[";
            oleCommandBuilder2.QuoteSuffix = "]";
            button4_Click(null, null);

            this.Text = Main.idToAssetNo(Int32.Parse(id)) + ", " + id + " - View Details";
        }

        private void ViewDetails_Load(object sender, EventArgs e)
        {

        }

        BindingSource bindingSource;
        public void button3_Click(object sender, EventArgs e)
        {
            dataGridView1.EndEdit();
            if (da != null)
                da.Update(table);
            dataGridView1.DataSource = null;
            table.Clear();
            da.Fill(table);
            bindingSource = new BindingSource { DataSource = table };
            dataGridView1.DataSource = bindingSource;
            DataGridViewTextBoxColumn d = new DataGridViewTextBoxColumn();
            
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
            try
            {
                (table.Rows[dataGridView1.CurrentRow.Index])[dataGridView1.SelectedCells[0].ColumnIndex] = richTextBox1.Text;
            }catch(Exception er)
            {

            }
            //dataGridView1.Refresh();
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1 != null && dataGridView1.SelectedCells.Count > 0)
                richTextBox1.Text = dataGridView1.SelectedCells[0].Value.ToString();
        }
        BindingSource bindingSource2;
        public void button4_Click(object sender, EventArgs e)
        {
            dataGridView2.EndEdit();
            if (da2 != null)
                da2.Update(table2);
            dataGridView2.DataSource = null;
            table2.Clear();
            da2.Fill(table2);
            bindingSource2 = new BindingSource { DataSource = table2 };
            dataGridView2.DataSource = bindingSource2;
            dataGridView2.Columns[2].Visible = false;
        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if(e.RowIndex >= 0)
            new ViewDetails(m, table2.Rows[e.RowIndex][e.ColumnIndex].ToString()).Show();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            new New_Event(this, id).Show();
        }

        private void richTextBox1_Leave(object sender, EventArgs e)
        {
            button3_Click(null, null);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            new LinkAsset(this, id).Show();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            new ReplaceAsset(this, id).Show();
        }

        private void dataGridView2_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
           if ((e.ColumnIndex == 0 || e.ColumnIndex == 1 ) && e.Value != null)
            {
                e.Value = Main.idToAssetNo(Int32.Parse(e.Value.ToString()));
               // e.FormattingApplied = true;
            }
        }

        private void dataGridView1_UserDeletingRow(object sender, DataGridViewRowCancelEventArgs e)
        {
            e.Cancel = MessageBox.Show("Do you want really to delete the selected row", "Confirm", MessageBoxButtons.OKCancel) != DialogResult.OK; ;
        }
    }
}
