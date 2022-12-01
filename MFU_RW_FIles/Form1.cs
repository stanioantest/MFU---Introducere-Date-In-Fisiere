using IronXL;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp9
{
    public partial class Form1 : Form
    {
        List<decimal> numere = new List<decimal>();

        public Form1()
        {
            InitializeComponent();
            textBox1.Text = "E:\\Roboti\\MFU\\02.DataLogFile_31.10.2022.xlsx";
            textBox2.Text = "E:\\Roboti\\MFU\\test.xlsx";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ReadExcelFile();
            WriteExcelFile();
        }

        public void ReadExcelFile()
        {
            WorkBook workbook = WorkBook.Load(textBox1.Text);
            WorkSheet sheet = workbook.WorkSheets.First();

            //This is how we get range from Excel worksheet
            var range = sheet["AVX2:AVX51"];
            //This is how we can iterate over our range and read or edit any cell
            foreach (var cell in range)
            {
                numere.Add(Convert.ToDecimal(cell.Value));
            }
        }

        public void WriteExcelFile()
        {
            WorkBook workbook2 = WorkBook.Load(textBox2.Text);
            WorkSheet sheet2 = workbook2.DefaultWorkSheet;

            for (int i = 0; i < numere.Count; i++)
            {
                sheet2["C" + (i + 5)].Value = Convert.ToDecimal(numere[i]);
                sheet2["C" + (i + 5)].First().FormatString = "0.000";
                //  sheet2["C" + (i + 5)].NumberFormat = "##.000";
            }
            //Save Changes
            workbook2.SaveAs(textBox2.Text);
        }
    }
}
