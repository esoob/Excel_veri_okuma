using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO.Ports;
using System.Threading;

using Excel = Microsoft.Office.Interop.Excel;

namespace excel_veri_okuma_example
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public double[] veriler = new double[128];

        private void button1_Click(object sender, EventArgs e)
        {
            Excel.Application exc = new Excel.Application();

            Excel.Workbook excelWorkbook = exc.Workbooks.Open("C:\\Users\\esmab\\Desktop\\excel_veri_okuma_example\\sinüs_verileri.xlsx", 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            Excel.Sheets excelSheets = excelWorkbook.Worksheets;
            string currentSheet = "Sayfa1";
            Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelSheets.get_Item(currentSheet);
            string cellValue = (string)(excelWorksheet.Cells[10, 2] as Excel.Range).Value.ToString();
            string[] veriler1 = new string[128];



            for (int i = 0; i <= 127; i++)
            {
                listBox1.Items.Add((string)(excelWorksheet.Cells[i + 1, 1] as Excel.Range).Value.ToString());
                listBox2.Items.Add((string)(excelWorksheet.Cells[i + 1, 2] as Excel.Range).Value.ToString());
                veriler[i] = Convert.ToDouble(listBox2.Items[i]);
            }

            timer1.Interval = 20;
            timer1.Enabled = true;
            timer1.Tick += new EventHandler(timer1_Tick);
            timer1.Start();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {

        }
    }
}
