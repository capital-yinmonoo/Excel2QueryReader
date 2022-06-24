using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Excel_Reader
{
    public partial class Form2 : Form
    {
        DataTable dtMain = new DataTable();
        public Form2(DataTable dt)
        {
            InitializeComponent();
            dtMain = dt;
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            dataGridView1.DataSource = dtMain;
        }
    }
}
