using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CKM_Percent_Calculator
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_KeyPress(object sender, KeyPressEventArgs e)
        {
            
                if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) &&
                    (e.KeyChar != '.'))
                {
                    e.Handled = true;
                }

                // only allow one decimal point
                if ((e.KeyChar == '.') && ((sender as TextBox).Text.IndexOf('.') > -1))
                {
                    e.Handled = true;
                }
             
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                textBox4.Text = (Convert.ToDouble(textBox5.Text) / Convert.ToDouble(textBox6.Text) * 100).ToString();
                label10.Visible = false;
            }
            catch
            {
                label10.Visible = true;
            }

            try
            {
                textBox3.Text = (Convert.ToDouble(textBox2.Text) * (Convert.ToDouble(textBox1.Text) / 100)).ToString();
                label9.Visible = false;
            }
            catch
            {
                label9.Visible = true;
            }
        }
    }
}
