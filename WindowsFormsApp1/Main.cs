using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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

        public String connParam = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\shihao\Desktop\Database11.accdb;Persist Security Info=False;";

        static Main m;

        public static string idToAssetNo(int id)
        {
            for (int i = 0; i < m.table.Rows.Count; i++)
            {
                if (!(m.table.Rows[i].RowState == DataRowState.Deleted))
                    if (Int32.Parse(m.table.Rows[i]["ID"].ToString()) == id)
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
            connParam = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Persist Security Info=False;";
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
            dataGridView1.SelectionMode = DataGridViewSelectionMode.RowHeaderSelect;

            dataGridView1.DataSource = null;
            table.Clear();
            da.Fill(table);
            dataGridView1.DataSource = bindingSource;
            button2_Click(null, null);
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
            if (da != null)
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
            search = "";
            string[] seps = { " & " };
            string[] x = textBox1.Text.Split(seps, StringSplitOptions.RemoveEmptyEntries);
            foreach (string t in x) {
                if (t.Contains(": "))
                {
                    t.Substring(t.IndexOf(": "));
                    if (table.Columns.Contains(t.Substring(0, t.IndexOf(": "))))
                        search += "AND ([" + t.Substring(0, t.IndexOf(": ")) + "] LIKE '" + t.Substring(t.IndexOf(": ") + 2) + "*')";

                }
                else
                {
                    search = "AND ([Asset Number] LIKE '" + textBox1.Text + "*')";
                }
            }
            System.Console.WriteLine(search);
            genRowFilter();
            dataGridView1.Refresh();
        }

        private void dataGridView1_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            new ViewDetails(this, dataGridView1.Rows[e.RowIndex].Cells["ID"].Value.ToString()).Show();
        }
        string surp = "True";

        private void genRowFilter()
        {
            table.DefaultView.RowFilter = "(" + surp.ToString() + ") " + search.ToString() + "";
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
        iTextSharp.text.Font edge = FontFactory.GetFont("Arial", 8, iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
        iTextSharp.text.Font Title = FontFactory.GetFont("Arial", 18, iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
        iTextSharp.text.Font content = FontFactory.GetFont("Arial", 5, iTextSharp.text.Font.NORMAL, BaseColor.BLACK);

        public PdfPCell getCell(String text, int alignment, iTextSharp.text.Font f)
        {
            Phrase p = new Phrase(text, f);
            PdfPCell cell = new PdfPCell(p);
            cell.UseBorderPadding = false;
            cell.Border = iTextSharp.text.Rectangle.NO_BORDER;
            cell.HorizontalAlignment = alignment;
            return cell;
        }

        public PdfPTable generateTable(DataGridViewSelectedRowCollection stuff, iTextSharp.text.Font f)
        {
            PdfPTable ret = new PdfPTable(5);
            ret.WidthPercentage = 100;
            ret.AddCell("Asset No");
            ret.AddCell("Model");
            ret.AddCell("Serial No");
            ret.AddCell("HDD");
            ret.AddCell("Memory");
            foreach (DataGridViewRow i in stuff)
            {
                DataRow b = ((DataRowView)i.DataBoundItem).Row;
                ret.AddCell(new Phrase(b["Asset Number"].ToString(), f));
                ret.AddCell(new Phrase(b["Model"].ToString(), f));
                ret.AddCell(new Phrase(b["S/N"].ToString(), f));
                ret.AddCell(new Phrase(b["HDD (GB)"].ToString() + " GB", f));
                ret.AddCell(new Phrase(b["Memory (GB)"].ToString() + " GB", f));
            }
            return ret;
        }

        public void refresh2()
        {
            table.Clear();
            da.Fill(table);
            this.Refresh();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //Generate Report
            FileStream myStream;
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();

            saveFileDialog1.Filter = "pdf files (*.pdf)|*.pdf|All files (*.*)|*.*";
            saveFileDialog1.FilterIndex = 1;
            saveFileDialog1.RestoreDirectory = true;

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                if ((myStream = (FileStream)saveFileDialog1.OpenFile()) != null)
                {
                    // Code to write the stream goes here.
                    Document doc = new Document(PageSize.A4,36,36,36,36);
                    PdfWriter writer = PdfWriter.GetInstance(doc, myStream);
                    doc.Open();
                    PdfPTable tab = new PdfPTable(3);
                    tab.WidthPercentage = 100;
                    tab.DefaultCell.Border = iTextSharp.text.Rectangle.NO_BORDER;
                    tab.AddCell(getCell("174 Stone Road W", PdfPCell.ALIGN_LEFT,edge));
                    tab.AddCell(getCell("Asset Report", PdfPCell.ALIGN_CENTER, Title));
                    tab.AddCell(getCell(DateTime.Now.ToString("yyyy-MM-dd"), PdfPCell.ALIGN_RIGHT, edge));
                    tab.SpacingAfter = 36;
                    doc.Add(tab);
                    doc.Add(generateTable(dataGridView1.SelectedRows, content));

                    doc.Close();
                    System.Diagnostics.Process.Start(saveFileDialog1.FileName);
                }
            }

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            surp = "[Status] LIKE '" + comboBox1.SelectedText + "*'";
            genRowFilter();
            dataGridView1.Refresh();
        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            foreach (DataGridViewCell a in dataGridView1.SelectedCells)
            {
                table.Rows[a.RowIndex][a.ColumnIndex] = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            newGroup d = new newGroup(this);
            if(d.ShowDialog() == DialogResult.OK)
            {
                if (conn.State == ConnectionState.Closed)
                    conn.Open();
                OleDbCommand c = new OleDbCommand("INSERT INTO [GroupLink] ([GID], [ID]) Values (@a,@b)", conn);
                c.Parameters.AddWithValue("@a", d.ID);
                c.Parameters.AddWithValue("@b", 0);
                foreach ( DataGridViewRow r in dataGridView1.SelectedCells.Cast<DataGridViewCell>()
                                           .Select(cell => cell.OwningRow)
                                           .Distinct())
                {
                    Console.Out.WriteLine("test");
                    c.Parameters[0].Value = d.ID;
                    c.Parameters[1].Value = ((DataRowView)r.DataBoundItem).Row["ID"];
                    c.ExecuteNonQuery();

                }
                button2_Click(null,null);
            }
            
        }

        private void button5_Click(object sender, EventArgs e)
        {
            new GroupManager(this).Show();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            GroupSelect d = new GroupSelect();
            if (d.ShowDialog() == DialogResult.OK)
            {
                if (conn.State == ConnectionState.Closed)
                    conn.Open();
                OleDbCommand c = new OleDbCommand("INSERT INTO [GroupLink] ([GID], [ID]) Values (@a,@b)", conn);
                c.Parameters.AddWithValue("@a", d.GID);
                c.Parameters.AddWithValue("@b", 0);
                foreach (DataGridViewRow r in dataGridView1.SelectedCells.Cast<DataGridViewCell>()
                                           .Select(cell => cell.OwningRow)
                                           .Distinct())
                {
                    Console.Out.WriteLine("test");
                    c.Parameters[0].Value = d.GID;
                    c.Parameters[1].Value = ((DataRowView)r.DataBoundItem).Row["ID"];
                    c.ExecuteNonQuery();

                }
                button2_Click(null, null);
            }
        }
    }
}
