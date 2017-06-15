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
    public partial class GroupManager : Form
    {

        OleDbCommand cmd;
        //parameter from mdsaputra.udl
        private OleDbCommandBuilder oleCommandBuilder = null;
        OleDbDataAdapter da;
        BindingSource bindingSource;
        public DataTable table = new DataTable();

        OleDbCommand cmd2;
        //parameter from mdsaputra.udl
        private OleDbCommandBuilder oleCommandBuilder2 = null;
        OleDbDataAdapter da2;
        BindingSource bindingSource2;
        public DataTable table2 = new DataTable();

        
        Main m;
        public GroupManager(Main refer)
        {
            InitializeComponent();
            m = refer;
            string strSQL = "SELECT * FROM [Groups]";  //rename Sheet$ to yours sheet name (Code$ you said)
            cmd = new OleDbCommand(strSQL, m.conn);
            da = new OleDbDataAdapter(cmd);
            oleCommandBuilder = new OleDbCommandBuilder(da);
            oleCommandBuilder.QuotePrefix = "[";
            oleCommandBuilder.QuoteSuffix = "]";
            bindingSource = new BindingSource { DataSource = table };
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridView1.MultiSelect = false;
            dataGridView1.DataSource = null;
            table.Clear();
            da.Fill(table);
            dataGridView1.DataSource = bindingSource;


            string strSQL2 = "SELECT * FROM [GEvents]";  //rename Sheet$ to yours sheet name (Code$ you said)
            cmd2 = new OleDbCommand(strSQL2, m.conn);
            da2 = new OleDbDataAdapter(cmd2);
            oleCommandBuilder2 = new OleDbCommandBuilder(da2);
            oleCommandBuilder2.QuotePrefix = "[";
            oleCommandBuilder2.QuoteSuffix = "]";
            bindingSource2 = new BindingSource { DataSource = table2 };
            dataGridView2.SelectionMode = DataGridViewSelectionMode.RowHeaderSelect;
            dataGridView1.MultiSelect = false;
            dataGridView2.DataSource = null;
            table2.Clear();
            da2.Fill(table2);
            dataGridView2.DataSource = bindingSource2;

            dataGridView3.SelectionMode = DataGridViewSelectionMode.FullRowSelect;



        }

        public void updateAssetTable()
        {
            try
            {
                using (var cn = new OleDbConnection(m.connParam))
                using (var cmd3 = new OleDbCommand("SELECT * FROM [Assets] WHERE [ID] IN (SELECT [ID] FROM [GroupLink] WHERE [GID] LIKE ?) "))
                {
                    // Parameter names don't matter; OleDb uses positional parameters.
                    if (dataGridView1.SelectedRows.Count > 0)
                    {
                        cn.Open();
                        cmd3.Connection = cn;
                        cmd3.Parameters.AddWithValue("@p0", dataGridView1.SelectedRows[0].Cells["GID"].Value.ToString());

                        Console.Out.WriteLine(dataGridView1.SelectedRows[0].Cells["GID"].Value.ToString());

                        var objDataSet = new DataTable();
                        var objDataAdapter = new OleDbDataAdapter(cmd3);
                        objDataAdapter.Fill(objDataSet);

                        dataGridView3.DataSource = objDataSet;
                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
                MessageBox.Show(ex.StackTrace.ToString());
            }
        }


        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            dataGridView1.EndEdit();
            da.Update(table);
            table.AcceptChanges();
            table.Clear();
            da.Fill(table);
        }

        public void button3_Click(object sender, EventArgs e)
        {
            dataGridView2.EndEdit();
            da2.Update(table2);
            table2.AcceptChanges();
            table2.Clear();
            da2.Fill(table2);
        }

        private void GroupManager_FormClosing(object sender, FormClosingEventArgs e)
        {
            if ((table.GetChanges() != null || table2.GetChanges() != null) && MessageBox.Show("Save Changes if any?", "Confirm", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                button3_Click(null, null);
                button2_Click(null, null);
            }
        }
        private void dataGridView3_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            
            updateAssetTable();
            if(dataGridView1.SelectedCells.Count > 0)
            table2.DefaultView.RowFilter = "[GID] = " + dataGridView1.SelectedCells[0].Value.ToString();
        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView2 != null && dataGridView2.SelectedCells.Count > 0)
                richTextBox1.Text = dataGridView2.SelectedCells[0].Value.ToString();
        }

        private void GroupManager_KeyDown(object sender, KeyEventArgs e)
        {
        }

        private void dataGridView1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == 46 && dataGridView1.SelectedRows.Count > 0)
                e.Handled = MessageBox.Show("Do you want really to delete the selected rows", "Confirm", MessageBoxButtons.OKCancel) != DialogResult.OK;
        }

        private void dataGridView2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == 46 && dataGridView2.SelectedRows.Count > 0)
                e.Handled = MessageBox.Show("Do you want really to delete the selected rows", "Confirm", MessageBoxButtons.OKCancel) != DialogResult.OK;
        }


        private void dataGridView3_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == 46 && dataGridView3.SelectedRows.Count > 0)
                e.Handled = MessageBox.Show("Do you want really to delete the selected rows", "Confirm", MessageBoxButtons.OKCancel) == DialogResult.OK;
            if (e.Handled)
            {
                try
                {
                    using (var cn = new OleDbConnection(m.connParam))
                    using (var cmd3 = new OleDbCommand("DELETE FROM [GroupLink] WHERE [GID] LIKE ? AND [ID] LIKE ?"))
                    {
                        cn.Open();
                        cmd3.Connection = cn;
                        cmd3.Parameters.AddWithValue("", -1);
                        cmd3.Parameters.AddWithValue("", -1);
                        // Parameter names don't matter; OleDb uses positional parameters.
                        if (dataGridView1.SelectedRows.Count > 0)
                        {
                            foreach (DataGridViewRow r in dataGridView3.SelectedRows)
                            {

                                cmd3.Parameters[0].Value = dataGridView1.SelectedRows[0].Cells["GID"].Value.ToString();
                                cmd3.Parameters[1].Value = r.Cells["ID"].Value.ToString();

                                 Console.Out.WriteLine(dataGridView1.SelectedRows[0].Cells["GID"].Value.ToString() + " ASDASD");
                                cmd3.ExecuteNonQuery();
                            }
                        }

                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString());
                    MessageBox.Show(ex.StackTrace.ToString());
                }
                updateAssetTable();
                
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if(dataGridView1.SelectedRows.Count > 0)
            new New_GEvent(this, dataGridView1.SelectedRows[0].Cells["GID"].Value.ToString()).Show();
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if ((table2.Rows[dataGridView2.CurrentRow.Index])[dataGridView2.SelectedCells[0].ColumnIndex].ToString() != richTextBox1.Text)
                    (table2.Rows[dataGridView2.CurrentRow.Index])[dataGridView2.SelectedCells[0].ColumnIndex] = richTextBox1.Text;
            }
            catch (Exception er)
            {

            }
        }

        private void dataGridView3_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
                new ViewDetails(m, dataGridView3.Rows[e.RowIndex].Cells["ID"].Value.ToString()).Show();
        }

        private void GroupManager_Load(object sender, EventArgs e)
        {

        }
    }
}
