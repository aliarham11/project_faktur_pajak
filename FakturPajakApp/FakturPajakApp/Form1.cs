using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Excel=Microsoft.Office.Interop.Excel;

namespace FakturPajakApp
{
    public partial class Form1 : Form
    {
        public string templatePath;
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
            fdlg.FilterIndex = 1;
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
            fdlg.FilterIndex = 1;
            fdlg.RestoreDirectory = true;
            if (fdlg.ShowDialog() == DialogResult.OK)
            {
                tbxPathTemplate.Text = fdlg.FileName;
                templatePath = fdlg.FileName;
            }
        }

        public void newFakturPajak()
        {
            string savePath= "c:\\";

            SaveFileDialog sfdlg = new SaveFileDialog();
            sfdlg.InitialDirectory = @"c:\";
            sfdlg.Filter = "MS excel files (*.xls)|*.xls|MS excel 2010++ files (*.xlsx)|*.xlsx";
            sfdlg.FilterIndex = 1;
            sfdlg.RestoreDirectory = true;
            if (sfdlg.ShowDialog() == DialogResult.OK)
            {
                savePath = sfdlg.FileName;
            }

            Excel.Workbook MyBook = null;
            Excel.Application MyApp = null;
            Excel.Worksheet MySheet = null;

            MyApp = new Excel.Application();
            MyApp.Visible = false;
            MyBook = MyApp.Workbooks.Open(templatePath);
            MySheet = (Excel.Worksheet)MyBook.Sheets[1];


            MySheet.Cells[4, 7] = tbxNoFP1.ToString() + "-" + tbxNoFP2.ToString();
            MySheet.Cells[11, 6] = tbxNamaPerusahaan.Text.ToString();
            MySheet.Cells[19, 4] = tbxDeskripsi.Lines.ElementAt(0).ToString();
            MySheet.Cells[20, 4] = tbxDeskripsi.Lines.ElementAt(1).ToString();
            MySheet.Cells[21, 4] = tbxDeskripsi.Lines.ElementAt(2).ToString();
            MySheet.Cells[22, 4] = tbxDeskripsi.Lines.ElementAt(3).ToString();
            MySheet.Cells[19, 9] = tbxNominal.Text.ToString();
            MySheet.Cells[36, 9] = tbxDPP.Text.ToString();
            MySheet.Cells[37, 9] = tbxPPN.Text.ToString();
            dateTimePicker1.Format = DateTimePickerFormat.Custom;
            dateTimePicker1.CustomFormat = "MM/dd/yyyy";

            MySheet.Cells[40, 10] = dateTimePicker1.ToString();
            try
            {
                MyBook.SaveAs(savePath);
                MyBook.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void tbxDPP_TextChanged(object sender, EventArgs e)
        {
            try
            {
                tbxPPN.Text = (Double.Parse(tbxDPP.Text.ToString()) * 0.1).ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnSubmit_Click(object sender, EventArgs e)
        {
            newFakturPajak();
        }
    }
}
