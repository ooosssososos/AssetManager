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
using System.Management;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Security.Permissions;
using System.Net.Sockets;
using System.Net;
using System.DirectoryServices.AccountManagement;
using System.Reflection;

namespace WindowsFormsApp1
{
    public partial class Main : Form
    {
        public OleDbConnection conn;
        public static OleDbConnection con;
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
            con = new OleDbConnection(connParam);
            InitializeComponent();
            m = this;
        }
        BindingSource bindingSource;
        private void Form1_Load(object sender, EventArgs e)
        {

            // dataGridView1.Dock = DockStyle.Fill;
            dataGridView1.MinimumSize = new Size(0, 0);
            dataGridView1.MaximumSize = new Size(0, 0);
            dataGridView1.Margin = new Padding(0);
            tableLayoutPanel1.Anchor = (AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top);
            dataGridView1.Anchor = (AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top) ;
           // flowLayoutPanel2.Dock = DockStyle.Fill;
            conn = new OleDbConnection(connParam);
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
            Console.WriteLine("WAT");
            button2_Click(null, null);
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            if (!System.Windows.Forms.SystemInformation.TerminalServerSession)
            {
                Type dgvType = dataGridView1.GetType();
                PropertyInfo pi = dgvType.GetProperty("DoubleBuffered",
                  BindingFlags.Instance | BindingFlags.NonPublic);
                pi.SetValue(dataGridView1, true, null);
            }
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
            table.Clear();
            da.Fill(table);
            string strSQL2 = @"Select * from assets where ID In (
Select B.asset2 FROM assets A Left Join relations B ON A.ID = B.asset1 WHERE a.type = 'instrument' AND not B.asset1 IS NULL)
UNION
Select* from assets where ID In(
Select B.asset1 FROM assets A Left Join relations B ON A.ID = B.asset2 WHERE a.type = 'instrument' AND not B.asset1 IS NULL)";  //rename Sheet$ to yours sheet name (Code$ you said)
            OleDbCommand cmd2 = new OleDbCommand(strSQL2, m.conn);
            try { conn.Open(); } catch (Exception er) { }
            OleDbDataReader reader = cmd2.ExecuteReader();
            List<string> ins = new List<string>();
            while (reader.Read())
            {
                ins.Add(reader[0].ToString());
            }
            reader.Close();
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                try
                {
                    if (ins.Contains(row.Cells[0].Value.ToString()))
                    {
                        row.DefaultCellStyle.BackColor = Color.Aquamarine;
                    }
                }
                catch (Exception) { }
            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
        }

        private void bindingSource1_CurrentChanged(object sender, EventArgs e)
        {

        }
        string search = "";
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
        string surp = "true";

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
            if (e.KeyValue == 46 && dataGridView1.SelectedRows.Count > 0)
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
            surp = "[Status] LIKE '" + comboBox1.Text + "*'";
            if (comboBox1.Text == "")
                surp = "true";
            genRowFilter();
            dataGridView1.Refresh();
        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            string rep = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
            foreach (DataGridViewCell a in dataGridView1.SelectedCells)
            {
                ((DataRowView)a.OwningRow.DataBoundItem).Row[a.ColumnIndex] = rep;
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

        public static bool PingHost(string _HostURI, int _PortNumber)
        {
            try
            {
                var client = new TcpClient();
                var result = client.BeginConnect(_HostURI, _PortNumber, null, null);

                bool success = result.AsyncWaitHandle.WaitOne(TimeSpan.FromSeconds(3));

              //  Console.WriteLine(success);
               // Console.WriteLine("!"+!success);
                if (!success)
                {
                    return false;
                }

                // we have connected
               // Console.WriteLine("test3");
                client.EndConnect(result);
                //client.EndConnect(result);
               // Console.WriteLine("test2");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                return true;
            }
        }


        private void button7_Click(object sender, EventArgs e)
        {


            string[] login =  PasswordPrompt.ShowDialog();
            if (login == null)
            {
                MessageBox.Show("Empty Login", "Login Failed",
    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            bool valid = false;
            using (PrincipalContext context = new PrincipalContext(ContextType.Domain))
            {
                valid = context.ValidateCredentials(login[0], login[1]);
            }
            if (!valid)
            {
                MessageBox.Show("Incorrect Credentials", "Login Failed",
MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //Console.WriteLine(login[0]);
            ConnectionOptions options = new ConnectionOptions();
            options.EnablePrivileges = true;
            options.Impersonation = ImpersonationLevel.Impersonate;
            //   options.Authority = "ntlmdomain:" + System.Configuration.ConfigurationSettings.AppSettings["DomainName"].ToString();
            options.Authentication = AuthenticationLevel.Packet;
            options.Username = login[0];
            options.Password = login[1];

            EnumerationOptions opt = new EnumerationOptions();
            opt.Timeout = new TimeSpan(0, 0, 1);
            options.Timeout = new TimeSpan(0, 0, 1);
            List<Task> tasks = new List<Task>();
            foreach (DataRow r in table.Rows)
            {
                Console.WriteLine(DateTime.Now.Ticks / TimeSpan.TicksPerMillisecond);
              //  Console.WriteLine("Test1");
                if ((string)r["Asset Number"] == "") continue ;
                string name = (string)r["Asset Number"];
                if (String.Compare(((string)r["Type"]), "Laptop", StringComparison.OrdinalIgnoreCase) == 0) {
                    name += "P";
                }
                else if(String.Compare(((string)r["Type"]), "Desktop", StringComparison.OrdinalIgnoreCase) == 0)
                {
                    name += "D";
                }
                else
                {
                    continue;
                }

              //  Console.WriteLine("Testb" + name);
                tasks.Add(Task.Factory.StartNew(() =>
                {
                    try
                    {
                     //   if (!PingHost("ONGUELA" + name, 135))
                    //        return;

                        Console.WriteLine("test");
                        ManagementScope scope = new ManagementScope("\\\\ONGUELA" + name + "\\root\\cimv2", options);

                        scope.Options.Timeout = TimeSpan.FromSeconds(5);
                        scope.Connect();
                        ObjectQuery query = new ObjectQuery("SELECT Capacity FROM Win32_PhysicalMemory");
                        ManagementObjectSearcher searcher = new ManagementObjectSearcher(scope, query, opt);
                        searcher.Options.Timeout = TimeSpan.FromSeconds(5);
                        ManagementObjectCollection queryCollection = searcher.Get();

                        UInt64 Capacity = 0;
                        foreach (ManagementObject m in queryCollection)
                        {
                            Capacity += (UInt64)m["Capacity"];
                        }



                        //Console.WriteLine("hia");
                        lock (r)
                        {
                            r["Memory (GB)"] = Math.Round((double)Capacity / (1024 * 1024 * 1024));
                            r["LastWMIC"] = DateTime.Now.ToString("yyyy-MM-dd");
                        }
                        //Console.WriteLine("hic");
                    } catch(Exception er)
                    {
                        Console.WriteLine(er.ToString());
                    }
                }));
            }

            Task.WaitAll(tasks.ToArray());
            dataGridView1.Refresh();
        }

        private void flowLayoutPanel2_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}
