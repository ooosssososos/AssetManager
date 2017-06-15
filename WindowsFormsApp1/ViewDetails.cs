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
        public DataGridView datag;
        OleDbCommand cmd;
        OleDbCommand cmd2;
        //parameter from mdsaputra.udl
        private OleDbCommandBuilder oleCommandBuilder = null;
        OleDbDataAdapter da;
        public DataTable table = new DataTable();
        private OleDbCommandBuilder oleCommandBuilder2 = null;
        OleDbDataAdapter da2;
        public DataTable table2 = new DataTable();
        
        String id;
        public Main m;
        bool loaded = false;
        public ViewDetails(Main refer, String ID)
        {

            InitializeComponent();
            id = ID;
            datag = dataGridView1;
            m = refer;
                        string strSQL = "SELECT * FROM History WHERE id = " + id + "";  //rename Sheet$ to yours sheet name (Code$ you said)
                        cmd = new OleDbCommand(strSQL, m.conn);
                        da = new OleDbDataAdapter(cmd);
                        oleCommandBuilder = new OleDbCommandBuilder(da);
                        oleCommandBuilder.QuotePrefix = "[";
                        oleCommandBuilder.QuoteSuffix = "]";
            dataGridView1.DataSource = null;
            table.Clear();
            da.Fill(table);
            bindingSource = new BindingSource { DataSource = table };
            dataGridView1.DataSource = bindingSource;

            string strSQL2 = "SELECT * FROM relations WHERE asset1 = " + id + " OR asset2 = " + id+"";  //rename Sheet$ to yours sheet name (Code$ you said)
            cmd2 = new OleDbCommand(strSQL2, m.conn);
            da2 = new OleDbDataAdapter(cmd2);
            oleCommandBuilder2 = new OleDbCommandBuilder(da2);
            oleCommandBuilder2.QuotePrefix = "[";
            oleCommandBuilder2.QuoteSuffix = "]";

            dataGridView2.DataSource = null;
            table2.Clear();
            da2.Fill(table2);
            bindingSource2 = new BindingSource { DataSource = table2 };
            dataGridView2.DataSource = bindingSource2;
            dataGridView2.Columns[2].Visible = false;
            dataGridView2.Columns.Add("Type", "Type");

            this.Text = Main.idToAssetNo(Int32.Parse(id)) + ", " + id + " - View Details";
        }

        public void refresh2()
        {
            table2.Clear();
            da2.Fill(table2);
            table.Clear();
            da.Fill(table);
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
            table.AcceptChanges();
            table.Clear();
            da.Fill(table);

            
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
                try
                {
                if((table.Rows[dataGridView1.CurrentRow.Index])[dataGridView1.SelectedCells[0].ColumnIndex].ToString() != richTextBox1.Text)
                    (table.Rows[dataGridView1.CurrentRow.Index])[dataGridView1.SelectedCells[0].ColumnIndex] = richTextBox1.Text;
                }
                catch (Exception er)
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
            table2.AcceptChanges();

            table2.Clear();
            da2.Fill(table2);
        }
        int notused = 0;

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if(e.RowIndex >= 0 && Int32.TryParse(dataGridView2.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString(),out notused))
            new ViewDetails(m, table2.Rows[e.RowIndex][e.ColumnIndex-1].ToString()).Show();
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
           if ((e.ColumnIndex == 1 || e.ColumnIndex == 2 ) && e.Value != null)
            {
                if (e.Value.ToString() == this.id)
                    e.Value = "THIS";
                else {
                    string dat = Main.idToAssetNo(Int32.Parse(e.Value.ToString()));
                    if(dat != "")
                    {
                        dataGridView2.Rows[e.RowIndex].Cells["Type"].Value = m.findById(Int32.Parse(e.Value.ToString()))["type"].ToString();
                        e.Value = dat;
                    }
                    else
                    {
                        e.Value = "REMOVED";
                    }
                }
               // e.FormattingApplied = true;
            }
        }

        private void dataGridView1_UserDeletingRow(object sender, DataGridViewRowCancelEventArgs e)
        {
        }
        private void dataGridView1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == 46 && dataGridView1.SelectedRows.Count > 0)
                e.Handled = MessageBox.Show("Do you want really to delete the selected rows", "Confirm", MessageBoxButtons.OKCancel) != DialogResult.OK;
        }

        private void ViewDetails_FormClosing(object sender, FormClosingEventArgs e)
        {
            if((table.GetChanges() != null || table2.GetChanges() != null) && MessageBox.Show("Save Changes if any?", "Confirm", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                button3_Click(null, null);
                button4_Click(null, null);
            }
        }

        private void dataGridView2_CellContentClick_1(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
