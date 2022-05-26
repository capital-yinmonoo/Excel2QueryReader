using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data;
using System.Data.OleDb;

namespace Excel_Reader
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            dtRight = new DataTable();
            dtRight.Columns.Add("No");
            dtRight.Columns.Add("TableName");
            dtRight.Columns.Add("SheetName");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog
            {
                InitialDirectory = @"Documents",
                Title = "Browse Excel Files",

                CheckFileExists = true,
                CheckPathExists = true,

                DefaultExt = "xls",
                Filter = "Excel files (*.xls)|*.xls",
                FilterIndex = 2,
                //RestoreDirectory = true, 
                //ReadOnlyChecked = true,
                //ShowReadOnly = true
            };
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = openFileDialog1.FileName;
                textBox2.Text = Path.GetDirectoryName(openFileDialog1.FileName);
                // ReadExcel(textBox1.Text,".xls");
            }
        }
        private void outputFile_Click(object sender, EventArgs e)
        {
            using (var fbd = new FolderBrowserDialog())
            {
                DialogResult result = fbd.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    textBox2.Text = fbd.SelectedPath;
                }
            }

        }
        private void button3_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            //CheckForIllegalCrossThreadCalls = true;
            //progressBar2.Visible = true;
         
            if (string.IsNullOrWhiteSpace(textBox1.Text))
            {
                
                MessageBox.Show("Please choose a input file to generate query!");
                textBox1.Focus();
                this.Cursor = Cursors.Default; ;
                return;
            }
            var resultList = ReadExcel(textBox1.Text, ".xls", true);
            var dv = resultList.DefaultView;
            dv.Sort = "No";

            resultList = dv.ToTable();
            if (resultList.Rows.Count > 0)
            {
                dtRes = resultList;
                dataGridView1.DataSource = resultList;
                dataGridView1.Columns[2].Visible = checkBox2.Checked;
                dataGridView1.RefreshEdit();
                dataGridView1.Refresh();
            }
            this.Cursor = Cursors.Default; ;
            //progressBar2.Visible = false;
        }
        private   async void Convert1()
        {
            if (string.IsNullOrWhiteSpace(textBox1.Text))
            {
                MessageBox.Show("Please choose a input file to generate query!");
                textBox1.Focus();
                return;
            }
            var resultList = ReadExcel(textBox1.Text, ".xls", true);
            var dv = resultList.DefaultView;
            dv.Sort = "No";
            resultList = dv.ToTable();
            if (resultList.Rows.Count > 0)
            {
                dtRes = resultList;
                dataGridView1.DataSource = resultList;
                dataGridView1.Columns[2].Visible = checkBox2.Checked;
                dataGridView1.RefreshEdit();
                dataGridView1.Refresh();
            }
        }
        private string DropIfExist(string tableName)
        {
            string drop = "IF EXISTS(SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[" + tableName + "]') AND type in (N'U'))" + Environment.NewLine + "DROP TABLE[dbo].[" + tableName + "]" + Environment.NewLine + "GO";
            return drop;
        }
        private string Commentary(string tableName)
        {
            DateTime date = DateTime.Now;
            string pcName = Environment.MachineName;
            string formattedDate = date.ToString("yyyy-MM-dd HH:mm:ss");
            string formattednormalDate = date.ToString("yyyy/MM/dd");
            string cmt = "/*" + Environment.NewLine +
                        "*********************************************************************" + Environment.NewLine +
                        "** 名前　:契約コースマスタの作成スクリプト" + Environment.NewLine +
                        "** 機能　:[" + tableName + "]を作成する" + Environment.NewLine +
                        "** 作成日:" + formattedDate + Environment.NewLine +
                        "** 作成者:" + pcName + Environment.NewLine +
                        "*********************************************************************" + Environment.NewLine +
                        "** 修正・変更の履歴" + Environment.NewLine +
                        "** 日付         作業者    作業ID       作業内容" + Environment.NewLine +
                        "** ------------------------------------------------------------------" + Environment.NewLine +
                        "** " + formattednormalDate + "   氏名      XXXXXXXXXX   改定内容の概要" + Environment.NewLine +
                        "*********************************************************************" + Environment.NewLine +
                        "*/";
            return cmt;
        }
        private string ExtendedQuery(string tableName)
        {
            string execute =Environment.NewLine+ " EXEC sys.sp_addextendedproperty" + Environment.NewLine +
                "@name = N'Obic7fxVersion',@value = N'1.0.20220419.1',@level0type = N'SCHEMA'," + Environment.NewLine +
             "@level0name = N'dbo',@level1type = N'TABLE',@level1name = N'" + tableName + "'" + Environment.NewLine +
             "GO";
            return execute;
        }
        private string SakuseiTable(DataTable dtExcel)
        {
            string query = "CREATE TABLE [" + dtExcel.Rows[0]["TBName"].ToString() + "]" + Environment.NewLine + "(";
            int u = 0;
            foreach (DataRow dr in dtExcel.Rows)
            {
                u++;
                var comma = "";
                if (u != 1)
                {
                    comma = ",";
                }
                var defaultVal = "";
                var tcount = "] ";
                if (!String.IsNullOrEmpty(dr["DefaultValue"].ToString()))
                {
                    defaultVal = "CONSTRAINT [DF_" + dtExcel.Rows[0]["TBName"].ToString() + "_" + dr["ItemName"].ToString() + "] DEFAULT ('" + dr["DefaultValue"].ToString() + "')";
                }
                string[] dtype = new string[] { "date", "datetime", "tinyint", "int", "image", "bigint", "bit", };
                if (!string.IsNullOrEmpty(dr["TypeCount"].ToString()))
                {
                    tcount = "] (" + dr["TypeCount"].ToString() + ") ";
                }
                query += comma + "[" + dr["ItemName"].ToString() + "]     [" + dr["DataType"].ToString() + tcount + (String.IsNullOrEmpty(dr["NullAllow"].ToString()) ? "NOT NULL " : "NULL ") + defaultVal + " --" + dr["JName"].ToString() + Environment.NewLine;
            }
            query += ") ON [PRIMARY]" + Environment.NewLine + "GO";
            return query;
        }
        private DataTable GetSource()
        {
            return null;
        }
        public DataTable ReadExcel(string fileName, string fileExt, bool IsList = false, string tbName=null)
        {

            string conn = string.Empty;
            DataTable dtexcel = new DataTable();
            if (fileExt.CompareTo(".xls") == 0)
                conn = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';"; //for below excel 2007  
            else
                conn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0;HDR=NO';"; //for above excel 2007  
            using (OleDbConnection con = new OleDbConnection(conn))
            {
                con.Open();
                try
                {
                 

                    //String[] excelSheets = new String[dt.Rows.Count];
                    int i = 0;
                    DataTable dtr = new DataTable();
                    dtr.Columns.Add("No");
                    dtr.Columns.Add("TableName");
                    dtr.Columns.Add("SheetName");
                    // Add the sheet name to the string array.

                    if (IsList)
                    {
                        DataTable dt = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                        if (dt == null)
                        {
                            return null;
                        }
                        int h = 0;
                        foreach (DataRow row in dt.Rows)
                        {
                            var dttemp = new DataTable();
                            // excelSheets[i] = row["TABLE_NAME"].ToString();
                            var name = row["TABLE_NAME"].ToString();
                            OleDbDataAdapter temp = new OleDbDataAdapter("select * from [" + name + "]", con);
                            temp.Fill(dttemp);
                            if (dttemp.Rows.Count != 0 && dttemp.Columns.Count > 12 && dttemp.Rows[0][12].ToString() == "PK")
                            {
                                h++;
                                dtr.Rows.Add(
                                    new object[]
                                    {
                                 h,  dttemp.Rows[3][0].ToString(), name
                                    }
                                    );
                            }
                            i++;
                        }
                        con.Close();
                        return dtr;
                    }
                   if (!string.IsNullOrEmpty(tbName))
                    {
                        OleDbDataAdapter oleAdpt = new OleDbDataAdapter("select * from ["+tbName+"]", con); //here we read data from sheet1  
                        oleAdpt.Fill(dtexcel); //fill excel data into dataTable 
                    } 
                }
                catch (Exception ex)
                {
                    var msg = ex.Message;
                }
                con.Close();
            }
            return ExcelTable(dtexcel);

        }
        private DataTable ExcelTable(DataTable dtexcel)
        {
            DataTable dtResult = new DataTable();
            dtResult.Columns.Add("No");
            dtResult.Columns.Add("JName");
            dtResult.Columns.Add("ItemName");
            dtResult.Columns.Add("DataType");
            dtResult.Columns.Add("TypeCount");
            dtResult.Columns.Add("DefaultValue");
            dtResult.Columns.Add("NullAllow");
            dtResult.Columns.Add("IDColumn");
            dtResult.Columns.Add("Constraint");
            dtResult.Columns.Add("Remarks");
            dtResult.Columns.Add("Reference"); 
            dtResult.Columns.Add("PKey");
            dtResult.Columns.Add("Index1");
            dtResult.Columns.Add("Index2");
            dtResult.Columns.Add("Index3");
            dtResult.Columns.Add("Index4");
            dtResult.Columns.Add("Index5");
            dtResult.Columns.Add("Index6");
            dtResult.Columns.Add("Index7");
            dtResult.Columns.Add("Index8");
            dtResult.Columns.Add("Index9");
            dtResult.Columns.Add("Index10");
            dtResult.Columns.Add("TBName");
            int h = 0;
            foreach (DataRow dr in dtexcel.Rows)
            {
                h++;
                if (h > 5)
                {
                    if (!string.IsNullOrEmpty(dr[0].ToString()))
                        dtResult.Rows.Add(
                            new object[] {
                        dr[0].ToString(),
                        dr["F2"].ToString(),
                        dr["F3"].ToString(),
                        dr["F4"].ToString().Replace("identity","int")  ,
                        //dr["F5"].ToString().Contains("→")? dr["F5"].ToString().Split('→').Last():dr["F5"].ToString(),
                        CountAll(dr,dr["F5"].ToString()),
                        dr[5].ToString(),
                        dr["F7"].ToString(),
                        dr["F8"].ToString(),
                        dr["F9"].ToString(),
                        dr["F10"].ToString(),
                        dr["F11"].ToString(),
                        dr[12].ToString(),
                        dr["F14"].ToString(),
                        dr["F15"].ToString(),
                        dr["F16"].ToString(),
                        dr["F17"].ToString(),
                        dr["F18"].ToString(),
                        dr["F19"].ToString(),
                        dr["F20"].ToString(),
                        dr["F21"].ToString(),
                        dr["F22"].ToString(),
                        dr["F23"].ToString(),
                        dtexcel.Rows[3][0].ToString().Replace("\n","")
                            }
                             );
                }
            }
            return dtResult;
        }
        private string CountAll(DataRow dtype,string val)
        {
            var value = val;
           if (value.Contains("→") )
                value= value.Split('→').Last() ;
            if (value.Contains("."))
                value = value.Replace(".", ",");
            if (dtype["F4"].ToString().Contains("tiny") || dtype["F4"].ToString().Contains("int"))
                value = null;

            return value;
        }
        private string SakuseiPK(DataTable dtexcel)
        {
            string query = "";
            query += "ALTER TABLE [" + dtexcel.Rows[0]["TBName"].ToString() + "] ADD CONSTRAINT [PK_" + dtexcel.Rows[0]["TBName"].ToString() + "] PRIMARY KEY CLUSTERED" + Environment.NewLine + "(" + Environment.NewLine;
            bool isExistKey = false;
            foreach (DataRow dr in dtexcel.Rows)
            {
                if (!string.IsNullOrEmpty(dr["PKey"].ToString()))
                {
                    query += "[" + dr["ItemName"].ToString() + "] ASC,";
                    isExistKey = true;
                }
            }
            try
            {
                query = query.Remove(query.Length - 1);
            }
            catch { }
            if (!isExistKey)
                return "";
            query += Environment.NewLine;
            query += ") WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON" + Environment.NewLine;
            query += "[PRIMARY]" + Environment.NewLine + "GO";
            return query;
        }
        private string SakuseiIndexers(DataTable dt)
        {
            var query = "";
            for (int i = 1; i <= 10; i++)
            {
                query += SakuseiIndex(dt, i.ToString()) + Environment.NewLine;
            }
            return query;
        }
        private string SakuseiIndex(DataTable dtexcel, string index)
        {
            var query = "";
            var subq = "";
            var tableName = dtexcel.Rows[0]["TBName"].ToString();
            query += "CREATE NONCLUSTERED INDEX [IX_" + tableName + "_" + index.PadLeft(2, '0') + "] ON [" + tableName + "]" + Environment.NewLine + "(" + Environment.NewLine;
            bool isExistIndex = false;
            foreach (DataRow dr in dtexcel.Rows)
            {

                if (!string.IsNullOrEmpty(dr["Index" + Convert.ToInt32(index).ToString()].ToString().Trim()))
                {
                    subq += "[" + dr["ItemName"].ToString() + "] ASC, ";
                    isExistIndex = true;
                }
            }
            subq = subq.TrimEnd();
            if (!isExistIndex)
                return "";
            subq = subq.Remove(subq.Length - 1);
            query += subq + Environment.NewLine;
            //[CourseCD] ASC, [ExpStartDate] ASC, [ExpEndDate] ASC, [CourseName] ASC, [InsertOperator] ASC, [InsertIPAddress] ASC
            query += ") WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY] " + Environment.NewLine + "GO";

            return query;
        }
        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {

            if (dataGridView1.Rows.Count > 0)
            {
                dataGridView1.Columns[2].Visible = checkBox2.Checked;
                dataGridView1.RefreshEdit();
                dataGridView1.Refresh();
            }
            if (dataGridView2.Rows.Count > 0)
            {
                dataGridView2.Columns[2].Visible = checkBox2.Checked;
                dataGridView2.RefreshEdit();
                dataGridView2.Refresh();
            }
        }
        DataTable dtRight;
        private void button4_Click(object sender, EventArgs e)
        {
            if (dataGridView1.DataSource == null || dataGridView1.RowCount == 0)
                return;
            dtRight.Rows.Add(new object[] {
            dataGridView1.SelectedRows[0].Cells[0].Value,
            dataGridView1.SelectedRows[0].Cells[1].Value.ToString().Replace("\n",""),
            dataGridView1.SelectedRows[0].Cells[2].Value,

            });
            DataTable uniqueCols = dtRight.DefaultView.ToTable(true, new string[] { "No", "TableName", "SheetName" });
            var dv = uniqueCols.DefaultView;
            dv.Sort = "No";
            uniqueCols = dv.ToTable();
            dataGridView2.DataSource = uniqueCols;
            if (!checkBox2.Checked)
                dataGridView2.Columns[2].Visible = false;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (dataGridView2.DataSource == null || dataGridView2.RowCount == 0)
                return;
            if (dtRight.Rows.Count > 0)
            {
                var dv = dtRight.DefaultView;
                dv.Sort = "No";
                dtRight = dv.ToTable();
                string val = dataGridView2.SelectedRows[0].Cells[0].Value.ToString();
                dtRight = dtRight.DefaultView.ToTable(true, new string[] { "No", "TableName", "SheetName" }); 
                foreach (DataRow dr in dtRight.Select($"No='{val}'"))
                    dr.Delete();
                dtRight.AcceptChanges();
                //dataGridView2.DataSource = null;
                dataGridView2.DataSource = dtRight;

                if (!checkBox2.Checked)
                    dataGridView2.Columns[2].Visible = false;


                dataGridView2.Columns[0].Visible = false; 
                dataGridView2.RefreshEdit();
                dataGridView2.Refresh();
            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
          
        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        { 
                 
        }
        static DataTable dtRes = new DataTable();
        private void textBox3_TextChanged(object sender, EventArgs e)
        {  
            var dt =dataGridView1.DataSource as DataTable;
           
            var txt = textBox3.Text.TrimEnd();
            var condi = "[No] like '%" + txt + "%' or [TableName] like '%" + txt + "%' or [SheetName] like '%" + txt + "%' and 1=1";
            var condi2= " 1 = 1";
          DataRow[] drPaytable = dtRes.Select(condi);
            // DataRow[] dr2 = dt.Select(condi2);
            if (drPaytable != null && drPaytable.Count() > 0 && !string.IsNullOrWhiteSpace(txt))
            {
                dataGridView1.DataSource = drPaytable.CopyToDataTable();

                dataGridView1.Columns[0].Visible = false;
                if (!checkBox2.Checked)
                    dataGridView1.Columns[2].Visible = false;

            }
            else if (drPaytable.Count() == 0)
            {
                DataTable dt2 = new DataTable();
                dt2.Columns.Add("No");
                dt2.Columns.Add("TableName");
                dt2.Columns.Add("SheetName");

                dataGridView1.DataSource = dt2;
            }
            else
            {
                dataGridView1.DataSource = dtRes;
                dataGridView1.Columns[0].Visible = false;
                if (!checkBox2.Checked)
                    dataGridView1.Columns[2].Visible = false;
            }
      

        }

        private void button6_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            if (string.IsNullOrEmpty(textBox1.Text) )
            {
                MessageBox.Show("Please choose a input file to generate query!");
                textBox1.Focus();
                this.Cursor = Cursors.Default; 
                return;
            }
            if (string.IsNullOrEmpty(textBox2.Text))
            {
                MessageBox.Show("Please select output directory!");
                textBox2.Focus();
                this.Cursor = Cursors.Default; 
                return;
            }
           
            var dtresult = dataGridView2.DataSource as DataTable;
            //if (dtresult == null)
            //{
            //    MessageBox.Show("Please select input file!"); 
            //    this.Cursor = Cursors.Default;
            //    return;
            //}
            if (dtresult != null && dtresult.Rows.Count > 0)
            {
                foreach (DataRow dr in dtresult.Rows)
                {
                    var dt = ReadExcel(textBox1.Text, ".xls", false, dr[2].ToString());
                    var tbName = dr[1].ToString();
                    var query = "";
                    if (checkBox1.Checked)
                        query += DropIfExist(tbName) + Environment.NewLine;
                    query += Commentary(tbName) + Environment.NewLine;
                    query += SakuseiTable(dt) + Environment.NewLine;
                    query += SakuseiPK(dt) + Environment.NewLine;
                    query += SakuseiIndexers(dt) + Environment.NewLine;
                    query += ExtendedQuery(tbName) + Environment.NewLine;

                    if (!Directory.Exists(textBox2.Text + "\\Tables"))
                    {
                        Directory.CreateDirectory(textBox2.Text + "\\Tables");
                    }
                    // \r\n
                    query = query.Replace("\r\n\r\n\r\n", "");
                    if (radioButton1.Checked)
                    {
                        File.WriteAllText(textBox2.Text + "\\Tables\\" + tbName + ".sql", query, Encoding.GetEncoding(932));
                    }
                    if (radioButton2.Checked)
                    {
                        File.WriteAllText(textBox2.Text + "\\Tables\\" + tbName + ".sql", query, Encoding.GetEncoding(51932)); ;
                    }
                    if (radioButton3.Checked)
                    {
                        File.WriteAllText(textBox2.Text + "\\Tables\\" + tbName.Replace("\n","") + ".sql", query, Encoding.UTF8);
                    }
                }

                System.Diagnostics.Process.Start(textBox2.Text + "\\Tables");
            }
            else
            {
                MessageBox.Show("Please select table to convert query");
            }
            this.Cursor = Cursors.Default; ; 
        }
        private void Form1_Load(object sender, EventArgs e)
        { 

        } 
     
    }
}