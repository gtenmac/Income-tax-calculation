using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace 累進稅率
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            dataGridView1.AutoGenerateColumns = false;
            DataProcess.ReadFile();
            dataGridView1.DataSource = DataProcess._list;
            textBox1.TextChanged += TextBox1_TextChanged;
        }

        private void TextBox1_TextChanged(object sender, EventArgs e)
        {
            label1.Text = "";
            if (textBox1.Text == "")
                textBox1.Text = "0";
            label1.Text = DataProcess.ComputeValue(textBox1.Text);
        }
    }
}
