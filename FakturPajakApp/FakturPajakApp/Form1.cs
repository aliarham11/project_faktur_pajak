using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace FakturPajakApp
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            OpenFileDialog fdlg = new OpenFileDialog();
            fdlg.Title = "Browse Faktur Pajak";
            fdlg.InitialDirectory = @"c:\";
            fdlg.Filter = "MS excel files (*.xls)|*.xls|MS excel 2010++ files (*.xlsx)|*.xlsx";
            fdlg.FilterIndex = 2;
            fdlg.RestoreDirectory = true;
            if (fdlg.ShowDialog() == DialogResult.OK)
            {
                tbxPathReport.Text = fdlg.FileName;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog fdlg = new OpenFileDialog();
            fdlg.Title = "Browse Faktur Pajak";
            fdlg.InitialDirectory = @"c:\";
            fdlg.Filter = "MS excel files (*.xls)|*.xls|MS excel 2010++ files (*.xlsx)|*.xlsx";
            fdlg.FilterIndex = 2;
            fdlg.RestoreDirectory = true;
            if (fdlg.ShowDialog() == DialogResult.OK)
            {
                tbxPathTemplate.Text = fdlg.FileName;
            }
        }
    }
}
