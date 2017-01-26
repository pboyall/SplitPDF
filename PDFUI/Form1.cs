using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PDFUI
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void btnPDFs_Click(object sender, EventArgs e)
        {
            OpenFileDialog theDialog = new OpenFileDialog();
            theDialog.Title = "Open Rename File";
            theDialog.Filter = "XLSX files|*.xlsx";
            theDialog.InitialDirectory = txtFolders.Text;
            DialogResult result = theDialog.ShowDialog();
            if (result == DialogResult.OK) // Test result.
            {
                txtNames.Text = theDialog.FileName.ToString();
            }
        }
    }
}
