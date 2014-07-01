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
            //newFakturPajak();
            writeReport();
        }

        public string convertMonth(int num)
        {
            if (num == 1)
            {
                return "Januari";
            }

            if (num == 2)
            {
                return "Februari";
            }

            if (num == 3)
            {
                return "Maret";
            }

            if (num == 4)
            {
                return "April";
            }

            if (num == 5)
            {
                return "Mei";
            }

            if (num == 6)
            {
                return "Juni";
            }

            if (num == 7)
            {
                return "Juli";
            }

            if (num == 8)
            {
                return "Agustus";
            }

            if (num == 9)
            {
                return "September";
            }

            if (num == 10)
            {
                return "Oktober";
            }

            if (num == 11)
            {
                return "November";
            }

            if (num == 12)
            {
                return "Desember";
            }

            return "Unknown";
        }

        public void createSheet(string name, Excel.Workbook myWorkbook)
        {
            var xlSheets = myWorkbook.Sheets as Excel.Sheets;
            var newSheet = (Excel.Worksheet)xlSheets.Add(xlSheets[1], Type.Missing, Type.Missing, Type.Missing);
            newSheet.Name = name;
            //Creating Header
            newSheet.Cells[1, 1] = "Rekap Faktur Pajak Dan Nota Retur";
            newSheet.Cells[2, 1] = "Masa Pajak: "+ name +" 2014";
            newSheet.Cells[3, 1] = "No.";
            newSheet.Cells[4, 1] = "Urut";
            
            newSheet.Cells[3, 2] = "Faktur";
            newSheet.Cells[4, 2] = "Pajak";

            newSheet.Cells[3, 4] = "Tanggal";
            newSheet.Cells[4, 4] = "Faktur Pajak";

            newSheet.Cells[3, 5] = "NPWP";

            newSheet.Cells[3, 6] = "Nama";
            newSheet.Cells[4, 6] = "Pembeli";

            newSheet.Cells[3, 7] = "Nomor";
            newSheet.Cells[4, 7] = "DO/DN";

            newSheet.Cells[3, 8] = "Nama";
            newSheet.Cells[4, 8] = "Jenis/ No Invoice";

            newSheet.Cells[4, 9] = "Unit";

            newSheet.Cells[3, 10] = "DPP";
            newSheet.Cells[4, 10] = "Unit";

            newSheet.Cells[3, 11] = "DPP";
            newSheet.Cells[4, 11] = "Parts";

            newSheet.Cells[3, 12] = "DPP";
            newSheet.Cells[4, 12] = "Other";

            newSheet.Cells[3, 13] = "PPN";
            newSheet.Cells[4, 13] = "Unit";

            newSheet.Cells[3, 14] = "PPN";
            newSheet.Cells[4, 14] = "Parts";

            newSheet.Cells[3, 15] = "PPN";
            newSheet.Cells[4, 15] = "Other";

            newSheet.Cells[3, 16] = "PPNBM";

            newSheet.Cells[3, 17] = "Total";

            newSheet.Cells[3, 18] = "DPP";

            newSheet.Cells[3, 19] = "PPN";
            
        }

        public void writeReport()
        {

            Excel.Workbook myBook = null;
            Excel.Application myApp = null;
            Excel.Worksheet mySheet = null;
            try
            {
                bool isEmpty = false;
                bool found = false;
                string month;
                string pathReport;
                int lastRow;
                lastRow = 1;

                pathReport = tbxPathReport.Text.ToString();
                myApp = new Excel.Application();
                myApp.Visible = false;
                myBook = myApp.Workbooks.Open(pathReport);
                month = convertMonth((int)dateTimePicker1.Value.Month);
                Console.WriteLine(month);
                foreach (Excel.Worksheet sheet in myBook.Sheets)
                {
                    if (sheet.Name.Equals(month))
                    {
                        mySheet = (Excel.Worksheet)myBook.Sheets[month];
                        found = true;
                        break;
                    }

                }

                if (found == false)
                {
                    createSheet(month, myBook);
                    mySheet = (Excel.Worksheet)myBook.Sheets[month];
                    //create sheet
                }

                System.Array MyValues = (System.Array)mySheet.get_Range("A1", "A1000").Cells.Value;
                while (!isEmpty)
                {
                    if (MyValues.GetValue(lastRow, 1) == null)
                    {
                        //
                        isEmpty = true;
                    }
                    else lastRow++;

                }
                //lastRow = (int)mySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;

                //lastRow += 1;
                //MessageBox.Show(lastRow.ToString());

                dateTimePicker1.Format = DateTimePickerFormat.Custom;
                dateTimePicker1.CustomFormat = "MM/dd/yyyy";
                mySheet.Cells[lastRow, 1] = lastRow - 4;
                mySheet.Cells[lastRow, 2] = tbxNoFP1.Text;
                mySheet.Cells[lastRow, 3] = tbxNoFP2.Text;
                mySheet.Cells[lastRow, 4] = dateTimePicker1.Text;
                mySheet.Cells[lastRow, 6] = tbxNamaPerusahaan.Text;
                mySheet.Cells[lastRow, 18] = tbxDPP.Text;
                mySheet.Cells[lastRow, 19] = tbxPPN.Text;
                myBook.Save();
             //   myBook.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                myBook.Close();
            }

        }
    }
}
